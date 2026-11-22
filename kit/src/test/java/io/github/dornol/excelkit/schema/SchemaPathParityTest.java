package io.github.dornol.excelkit.schema;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.List;
import java.util.TreeSet;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A writer taken from an {@link ExcelKitSchema} is the same writer as one built by hand with the same columns:
 * same cells, same styles, same package, with every sheet-level option the writer has. A schema also reads
 * back what it wrote, across rolled-over sheets.
 */
class SchemaPathParityTest {

    public static class Row2 {
        String name;
        Integer qty;
        Double price;
        LocalDate day;
    }

    private static final ExcelKitSchema<Row2> SCHEMA = ExcelKitSchema.<Row2>builder()
            .column("Name", r -> r.name, (r, c) -> r.name = c.asString(), c -> c.bold(true))
            .column("Qty", r -> r.qty, (r, c) -> r.qty = c.asInt(), c -> c.type(ExcelDataType.INTEGER))
            .column("Price", r -> r.price, (r, c) -> r.price = c.asDouble(), c -> c.type(ExcelDataType.DOUBLE))
            .column("Day", r -> r.day, (r, c) -> r.day = c.asLocalDate())
            .build();

    private static List<Row2> rows(int n) {
        return IntStream.range(0, n).mapToObj(i -> {
            Row2 r = new Row2();
            r.name = i % 4 == 3 ? null : "name " + i;
            r.qty = i % 5 == 4 ? null : i;
            r.price = i * 1.25;
            r.day = LocalDate.of(2024, 1, 1).plusDays(i);
            return r;
        }).toList();
    }

    private static ExcelWriter<Row2> byHand() {
        return ExcelWriter.<Row2>create()
                .column("Name", r -> r.name, c -> c.bold(true))
                .column("Qty", r -> r.qty, c -> c.type(ExcelDataType.INTEGER))
                .column("Price", r -> r.price, c -> c.type(ExcelDataType.DOUBLE))
                .column("Day", r -> r.day);
    }

    private static ExcelWriter<Row2> withOptions(ExcelWriter<Row2> writer) {
        return writer.maxRows(4).table("SchemaTable").autoFilter(true).freezeRows(1)
                .summary(s -> s.sum("Qty").average("Price"));
    }

    /** Text, number format and font of every cell, the sheet names, and the parts of the package. */
    private static List<String> fingerprint(byte[] xlsx) throws Exception {
        List<String> out = new ArrayList<>();
        DataFormatter fmt = new DataFormatter();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            for (Sheet sheet : wb) {
                out.add("sheet " + sheet.getSheetName() + " tables=" + ((org.apache.poi.xssf.usermodel.XSSFSheet) sheet).getTables().size());
                for (Row row : sheet) {
                    for (Cell cell : row) {
                        XSSFCellStyle style = (XSSFCellStyle) cell.getCellStyle();
                        out.add(cell.getAddress() + " <" + fmt.formatCellValue(cell) + "> fmt=" + style.getDataFormatString()
                                + " bold=" + style.getFont().getBold());
                    }
                }
            }
        }
        out.add("parts " + new TreeSet<>(partNames(xlsx)));
        return out;
    }

    private static List<String> partNames(byte[] xlsx) throws Exception {
        java.nio.file.Path tmp = java.nio.file.Files.createTempFile("schema-parity", ".xlsx");
        try {
            java.nio.file.Files.write(tmp, xlsx);
            try (java.util.zip.ZipFile zip = new java.util.zip.ZipFile(tmp.toFile())) {
                return zip.stream().map(java.util.zip.ZipEntry::getName).toList();
            }
        } finally {
            java.nio.file.Files.deleteIfExists(tmp);
        }
    }

    @Test
    void aSchemaWriterWritesTheSameWorkbookAsOneBuiltByHand() throws Exception {
        ByteArrayOutputStream fromSchema = new ByteArrayOutputStream();
        withOptions(SCHEMA.excelWriter()).write(rows(10)).writeTo(fromSchema);
        ByteArrayOutputStream handMade = new ByteArrayOutputStream();
        withOptions(byHand()).write(rows(10)).writeTo(handMade);

        List<String> reference = fingerprint(handMade.toByteArray());
        assertTrue(reference.stream().anyMatch(l -> l.startsWith("sheet ") && l.contains("tables=1")), "the options took effect");
        assertTrue(reference.stream().anyMatch(l -> l.contains("<Sum>")) && reference.stream().anyMatch(l -> l.contains("<Average>")), "the summary was written");
        assertEquals(reference, fingerprint(fromSchema.toByteArray()));
    }

    @Test
    void theInitOptionsOfASchemaWriterChangeHowItStreamsNotWhatItWrites() throws Exception {
        ByteArrayOutputStream plain = new ByteArrayOutputStream();
        SCHEMA.excelWriter().write(rows(30)).writeTo(plain);
        ByteArrayOutputStream tuned = new ByteArrayOutputStream();
        SCHEMA.excelWriter(o -> o.rowAccessWindowSize(5).compressTempFiles(true).useSharedStrings(true))
                .write(rows(30)).writeTo(tuned);

        assertTrue(partNames(tuned.toByteArray()).contains("xl/sharedStrings.xml"), "shared strings were switched on");
        assertEquals(textOf(plain.toByteArray()), textOf(tuned.toByteArray()));
    }

    private static List<String> textOf(byte[] xlsx) throws Exception {
        List<String> out = new ArrayList<>();
        DataFormatter fmt = new DataFormatter();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            for (Sheet sheet : wb) for (Row row : sheet) for (Cell cell : row) {
                out.add(cell.getAddress() + "=" + fmt.formatCellValue(cell));
            }
        }
        return out;
    }

    @Test
    void aSchemaReadsBackWhatItWroteAcrossRolledOverSheets() throws Exception {
        List<Row2> written = rows(10);
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        SCHEMA.excelWriter().maxRows(4).write(written).writeTo(out);

        List<Row2> back = new ArrayList<>();
        for (int sheet = 0; sheet < 3; sheet++) {
            List<ReadResult<Row2>> results = new ArrayList<>();
            ExcelReader<Row2> reader = SCHEMA.excelReader(Row2::new).sheetIndex(sheet);
            reader.read(new ByteArrayInputStream(out.toByteArray()), results::add);
            for (ReadResult<Row2> r : results) {
                assertTrue(r.success(), String.valueOf(r.messages()));
                back.add(r.data());
            }
        }
        assertEquals(written.size(), back.size());
        for (int i = 0; i < written.size(); i++) {
            Row2 w = written.get(i), r = back.get(i);
            // null text comes back as empty or null, as for any reader
            if (w.name == null) assertTrue(r.name == null || r.name.isEmpty(), "row " + i); else assertEquals(w.name, r.name, "row " + i);
            if (w.qty == null) assertNull(r.qty, "row " + i); else assertEquals(w.qty, r.qty, "row " + i);
            assertEquals(w.price, r.price, "row " + i);
            assertEquals(w.day, r.day, "row " + i);
        }
    }

    @Test
    void aSchemaCsvWriterWritesTheSameBytesAsOneBuiltByHand() throws Exception {
        ByteArrayOutputStream fromSchema = new ByteArrayOutputStream();
        SCHEMA.csvWriter().write(rows(10)).writeTo(fromSchema);
        ByteArrayOutputStream handMade = new ByteArrayOutputStream();
        CsvWriter.<Row2>create()
                .column("Name", r -> r.name).column("Qty", r -> r.qty)
                .column("Price", r -> r.price).column("Day", r -> r.day)
                .write(rows(10)).writeTo(handMade);

        assertEquals(new String(handMade.toByteArray(), java.nio.charset.StandardCharsets.UTF_8),
                new String(fromSchema.toByteArray(), java.nio.charset.StandardCharsets.UTF_8));
    }
}
