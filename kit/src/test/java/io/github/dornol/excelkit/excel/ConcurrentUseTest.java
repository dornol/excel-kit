package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.concurrent.Callable;
import java.util.concurrent.CountDownLatch;
import java.util.concurrent.ExecutorService;
import java.util.concurrent.Executors;
import java.util.concurrent.Future;
import java.util.concurrent.TimeUnit;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Writers, workbooks, template writers and readers share no state other than a few JVM-wide POI settings, so any
 * number of them may run at once. Each thread writes and reads data that is unique to it, in a mix of writer kinds,
 * all started together; a row of one thread turning up in another's file, a failure, or a leftover temp file (the
 * leak guard) fails the test.
 */
class ConcurrentUseTest {

    private record Item(String owner, int n) {
    }

    private static List<Item> items(int owner, int rows) {
        return IntStream.range(0, rows).mapToObj(i -> new Item("owner-" + owner, i)).toList();
    }

    /** All data rows of all sheets, as "owner/n". */
    private static List<String> rowsOf(byte[] xlsx) throws Exception {
        List<String> out = new ArrayList<>();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            for (Sheet sheet : wb) {
                for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                    var row = sheet.getRow(r);
                    if (row == null || row.getCell(0) == null
                            || row.getCell(0).getCellType() != org.apache.poi.ss.usermodel.CellType.STRING
                            || !row.getCell(0).getStringCellValue().startsWith("owner-")) continue; // header and summary rows
                    out.add(row.getCell(0).getStringCellValue() + "/" + (int) row.getCell(1).getNumericCellValue());
                }
            }
        }
        return out;
    }

    private static List<String> expected(int owner, int rows) {
        return items(owner, rows).stream().map(i -> i.owner() + "/" + i.n()).toList();
    }

    private static byte[] decrypt(byte[] bytes, String password) throws Exception {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
            Decryptor dec = Decryptor.getInstance(new EncryptionInfo(fs));
            assertTrue(dec.verifyPassword(password));
            return dec.getDataStream(fs).readAllBytes();
        }
    }

    /** One round of work for one thread, cycling through the kinds of writer. */
    private static void work(int owner, int round) throws Exception {
        int rows = 20 + (owner * 7 + round * 3) % 40;
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        switch ((owner + round) % 5) {
            case 0 -> {
                ExcelWriter.<Item>create().column("Owner", Item::owner).column("N", Item::n).write(items(owner, rows)).writeTo(out);
                assertEquals(expected(owner, rows), rowsOf(out.toByteArray()));
            }
            case 1 -> {
                ExcelWriter.<Item>create().maxRows(7).table("Conc" + owner + "x" + round).column("Owner", Item::owner).column("N", Item::n)
                        .summary(s -> s.sum("N")).write(items(owner, rows)).writeTo(out);
                assertEquals(expected(owner, rows), rowsOf(out.toByteArray()));
            }
            case 2 -> {
                ExcelWriter.<Item>create().password("pw" + owner).column("Owner", Item::owner).column("N", Item::n)
                        .write(items(owner, rows)).writeTo(out);
                assertEquals(expected(owner, rows), rowsOf(decrypt(out.toByteArray(), "pw" + owner)));
            }
            case 3 -> {
                try (ExcelWorkbook book = ExcelWorkbook.create()) {
                    book.<Item>sheet("One").column("Owner", Item::owner).column("N", Item::n).write(items(owner, rows).stream());
                    book.<Item>sheet("Two").column("Owner", Item::owner).column("N", Item::n).write(items(owner, 3).stream());
                    book.finish().writeTo(out);
                }
                List<String> want = new ArrayList<>(expected(owner, rows));
                want.addAll(expected(owner, 3));
                assertEquals(want, rowsOf(out.toByteArray()));
            }
            default -> {
                ExcelWriter.<Item>create().column("Owner", Item::owner).column("N", Item::n).write(items(owner, rows)).writeTo(out);
                List<String> read = new ArrayList<>();
                ExcelReader.<Item>mapping(r -> new Item(r.get("Owner").asString(), r.get("N").asInt()))
                        .read(new ByteArrayInputStream(out.toByteArray()), (ReadResult<Item> res) -> {
                            assertTrue(res.success(), String.valueOf(res.messages()));
                            read.add(res.data().owner() + "/" + res.data().n());
                        });
                assertEquals(expected(owner, rows), read);
            }
        }
    }

    @Test
    void manyThreadsMixingWritersAndReadersDoNotDisturbEachOther() throws Exception {
        int threads = 12, rounds = 15;
        ExecutorService pool = Executors.newFixedThreadPool(threads);
        CountDownLatch start = new CountDownLatch(1);
        try {
            List<Future<Integer>> futures = new ArrayList<>();
            for (int t = 0; t < threads; t++) {
                final int owner = t;
                Callable<Integer> job = () -> {
                    start.await();
                    for (int round = 0; round < rounds; round++) work(owner, round);
                    return rounds;
                };
                futures.add(pool.submit(job));
            }
            start.countDown();
            for (Future<Integer> f : futures) assertEquals(rounds, f.get(120, TimeUnit.SECONDS));
        } finally {
            pool.shutdownNow();
        }
    }
}
