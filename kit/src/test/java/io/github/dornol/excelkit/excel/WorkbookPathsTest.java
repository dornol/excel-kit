package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeSet;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Two things the per-feature matrix cannot show: sheets of one workbook do not leak their features into each
 * other, and the ways of getting the bytes out (stream, file, password given late) all produce the same workbook.
 */
class WorkbookPathsTest {

    private record Item(String name, int qty) {
    }

    private static List<Item> items(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Item("item " + i, i)).toList();
    }

    // ---- sheets of one workbook ---------------------------------------------------------------------------

    @Test
    void eachSheetOfAWorkbookKeepsOnlyItsOwnFeatures() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("Rich").maxRows(3).table("Rich").protectSheet("pw").autoFilter(true).freezeRows(1)
                    .tabColor(ExcelColor.GREEN)
                    .column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s.label("Total").sum("Qty"))
                    .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"))
                    .conditionalFormatting(cf -> cf.columns(1).greaterThan("1", ExcelColor.LIGHT_RED))
                    .write(items(7).stream());
            book.<Item>sheet("Plain")
                    .column("Label", Item::name)
                    .write(items(4).stream());
            book.<Item>sheet("Grouped").table("Grouped2").maxRows(2)
                    .column("Name", Item::name).column("Qty", Item::qty)
                    .write(items(3).stream());
            book.finish().writeTo(out);
        }

        OoxmlStructureTest.Pkg pkg = new OoxmlStructureTest.Pkg(out.toByteArray());
        OoxmlStructureTest.assertValidPackage(pkg);
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            List<String> names = new ArrayList<>();
            for (Sheet s : wb) names.add(s.getSheetName());
            assertEquals(6, names.size(), names.toString()); // Rich x3 (7 rows / 3), Plain, Grouped x2 (3 rows / 2)
            assertEquals("Plain", names.get(3), names.toString());

            int plain = wb.getSheetIndex("Plain");
            String plainPart = "xl/worksheets/sheet" + (plain + 1) + ".xml";
            for (String feature : List.of("tablePart", "sheetProtection", "autoFilter", "pane", "drawing",
                    "conditionalFormatting", "tabColor")) {
                assertEquals(0, pkg.select(plainPart, OoxmlStructureTest.local(feature)).size(),
                        "'Plain' must not inherit " + feature + " from the first sheet");
            }
            // the plain sheet has no summary row either: its rows are the header and four items
            assertEquals(4, wb.getSheetAt(plain).getLastRowNum());
        }
        // Table names are unique across the whole workbook.
        List<String> tableNames = new ArrayList<>();
        for (String part : pkg.parts.keySet()) {
            if (part.matches("xl/tables/table\\d+\\.xml")) {
                tableNames.add(pkg.doc(part).getDocumentElement().getAttribute("name"));
            }
        }
        assertEquals(5, tableNames.size(), tableNames.toString());
        assertEquals(5, new TreeSet<>(tableNames).size(), "duplicate table names: " + tableNames);
    }

    @Test
    void theSameTableNameOnTwoSheetsIsRefusedNotWrittenAsABrokenFile() {
        // Excel refuses (or repairs) a workbook with two tables of one name, so it must not be written.
        assertThrows(ExcelWriteException.class, () -> {
            try (ExcelWorkbook book = ExcelWorkbook.create()) {
                book.<Item>sheet("A").table("Same").column("Name", Item::name).write(items(2).stream());
                book.<Item>sheet("B").table("Same").column("Name", Item::name).write(items(2).stream());
                book.finish().writeTo(java.io.OutputStream.nullOutputStream());
            }
        });
    }

    // ---- every way out gives the same workbook -------------------------------------------------------------

    private interface Sink {
        byte[] bytes(ExcelHandler handler, Path dir) throws Exception;
    }

    private static ExcelHandler handler(String password) {
        ExcelWriter<Item> writer = ExcelWriter.<Item>create().maxRows(4).table("PathsTable").autoFilter(true)
                .column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .summary(s -> s.label("Total").sum("Qty"));
        if (password != null) writer.password(password);
        return writer.write(items(10));
    }

    private static byte[] decrypt(byte[] bytes, String password) throws Exception {
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
            Decryptor dec = Decryptor.getInstance(new EncryptionInfo(fs));
            assertTrue(dec.verifyPassword(password));
            return dec.getDataStream(fs).readAllBytes();
        }
    }

    /** Every sheet name and every cell as the text a user sees, plus the set of package parts. */
    private static List<String> fingerprint(byte[] xlsx) throws Exception {
        List<String> out = new ArrayList<>();
        DataFormatter fmt = new DataFormatter();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            for (Sheet sheet : wb) {
                out.add("sheet " + sheet.getSheetName());
                for (Row row : sheet) {
                    StringBuilder line = new StringBuilder("  r" + row.getRowNum() + ":");
                    for (var cell : row) line.append(' ').append(fmt.formatCellValue(cell));
                    out.add(line.toString());
                }
            }
        }
        out.add("parts " + new TreeSet<>(new OoxmlStructureTest.Pkg(xlsx).parts.keySet()));
        return out;
    }

    @Test
    void streamFileAndLateBoundPasswordAllWriteTheSameWorkbook(@TempDir Path dir) throws Exception {
        ByteArrayOutputStream viaStream = new ByteArrayOutputStream();
        handler(null).writeTo(viaStream);
        List<String> reference = fingerprint(viaStream.toByteArray());
        assertTrue(reference.stream().anyMatch(l -> l.contains("Total")), "the reference has its summary rows");

        // file
        Path file = dir.resolve("plain.xlsx");
        ExcelHandler toFile = handler(null);
        toFile.writeTo(Files.newOutputStream(file));
        assertEquals(reference, fingerprint(Files.readAllBytes(file)), "plain file");

        // password bound by the writer, written to a stream
        ByteArrayOutputStream eager = new ByteArrayOutputStream();
        handler("pw").writeTo(eager);
        assertEquals(reference, fingerprint(decrypt(eager.toByteArray(), "pw")), "eager password");

        // password given when the response is written, to a stream and to a file
        ByteArrayOutputStream late = new ByteArrayOutputStream();
        handler(null).writeTo(late, "pw");
        assertEquals(reference, fingerprint(decrypt(late.toByteArray(), "pw")), "late password, stream");

        Path encrypted = dir.resolve("encrypted.xlsx");
        handler(null).writeTo(encrypted, "pw");
        assertEquals(reference, fingerprint(decrypt(Files.readAllBytes(encrypted), "pw")), "late password, file");

        // an encrypted file is never a plain workbook
        assertFalse(eager.toString(java.nio.charset.StandardCharsets.ISO_8859_1).startsWith("PK"));
        assertFalse(new String(Files.readAllBytes(encrypted), java.nio.charset.StandardCharsets.ISO_8859_1).startsWith("PK"));
    }
}
