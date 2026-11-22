package io.github.dornol.excelkit.excel;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;
import java.util.zip.ZipInputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The {@code .xlsx} a writer streams out is a ZIP that stream-based readers can read: {@code java.util.zip.ZipInputStream},
 * LibreOffice and similar tools read entry sizes from the local header or from a 4-byte data descriptor. POI 6 writes
 * every entry with an 8-byte (ZIP64) data descriptor by default, which those readers reject
 * ("invalid entry size (expected 0 but got ...)"; LibreOffice: "source file could not be loaded").
 */
class StreamingZipCompatibilityTest {

    private record Item(String name, int qty) {
    }

    private static List<Item> items(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Item("item " + i, i)).toList();
    }

    /** Reads every entry the way a streaming reader does; fails like such a reader on a descriptor it cannot parse. */
    private static List<String> entriesOf(byte[] xlsx) throws Exception {
        List<String> names = new ArrayList<>();
        try (ZipInputStream zip = new ZipInputStream(new ByteArrayInputStream(xlsx))) {
            for (var e = zip.getNextEntry(); e != null; e = zip.getNextEntry()) {
                assertTrue(zip.readAllBytes().length > 0, "entry " + e.getName() + " is empty");
                names.add(e.getName());
            }
        }
        return names;
    }

    private static void assertReadable(byte[] xlsx) throws Exception {
        List<String> names = entriesOf(xlsx);
        assertTrue(names.contains("[Content_Types].xml") && names.contains("xl/workbook.xml"), names.toString());
        // and POI itself still reads it
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            assertTrue(wb.getNumberOfSheets() >= 1);
        }
    }

    @Test
    void aWorkbookWrittenToAStreamIsReadableByAStreamingZipReader() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().column("Name", Item::name).column("Qty", Item::qty).write(items(20)).writeTo(out);
        assertReadable(out.toByteArray());
    }

    @Test
    void sheetsFromRolloverTablesAndChartsAreReadableToo() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().maxRows(5).table("ZipCompat").column("Name", Item::name).column("Qty", Item::qty)
                .write(items(20)).writeTo(out);
        assertReadable(out.toByteArray());

        ByteArrayOutputStream charted = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().column("Name", Item::name).column("Qty", Item::qty)
                .chart(c -> c.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"))
                .write(items(8)).writeTo(charted);
        assertReadable(charted.toByteArray());
    }

    @Test
    void aMultiSheetWorkbookIsReadable() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("A").column("Name", Item::name).write(items(3).stream());
            book.<Item>sheet("B").column("Qty", Item::qty).write(items(3).stream());
            book.finish().writeTo(out);
        }
        assertReadable(out.toByteArray());
    }

    @Test
    void aTemplateWriterOutputIsReadable() throws Exception {
        byte[] template;
        try (XSSFWorkbook wb = new XSSFWorkbook(); ByteArrayOutputStream o = new ByteArrayOutputStream()) {
            wb.createSheet("S").createRow(0).createCell(0).setCellValue("Title");
            wb.write(o);
            template = o.toByteArray();
        }
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.<Item>list(2).column("Name", Item::name).write(items(4));
            writer.finish().writeTo(out);
        }
        assertReadable(out.toByteArray());
    }

    @Test
    void theWorkbookInsideAnEncryptedFileIsReadable() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().password("pw").column("Name", Item::name).write(items(6)).writeTo(out);

        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(out.toByteArray()))) {
            Decryptor dec = Decryptor.getInstance(new EncryptionInfo(fs));
            assertTrue(dec.verifyPassword("pw"));
            assertReadable(dec.getDataStream(fs).readAllBytes());
        }
    }

    @Test
    void theEntriesKeepTheirNamesAndOrderForReadersThatExpectContentTypesFirst() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().column("Name", Item::name).write(items(2)).writeTo(out);
        assertEquals("[Content_Types].xml", entriesOf(out.toByteArray()).get(0));
    }
}
