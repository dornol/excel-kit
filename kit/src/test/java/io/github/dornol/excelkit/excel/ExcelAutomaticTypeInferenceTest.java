package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelHyperlink;
import io.github.dornol.excelkit.excel.feature.ExcelImage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Tests value-based type inference for column writers. */
class ExcelAutomaticTypeInferenceTest {

    @Test
    void mixedInferredValuesRetainTheirMeaningAndNumericType() throws IOException {
        var rows = List.<Map<String, Object>>of(
                Map.of("Numeric", new BigDecimal("1.25"), "Mixed", true),
                Map.of("Numeric", 2L, "Mixed", "yes"));
        try (var workbook = read(ExcelWriter.forMap("Numeric", "Mixed").write(rows))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("Y", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals("yes", sheet.getRow(2).getCell(1).getStringCellValue());
            assertEquals(2.0, sheet.getRow(2).getCell(0).getNumericCellValue());
        }
    }

    @Test
    void historicalDatesRemainDistinctAndTimesAreDayFractions() throws IOException {
        var values = List.of(LocalDate.of(1850, 1, 1), LocalDate.of(1899, 1, 1), LocalDate.of(1900, 1, 1));
        var writer = ExcelWriter.<LocalDate>create().writeErrorPolicy(ExcelWriteErrorPolicy.LENIENT).column("Date", value -> value)
                .column("Time", value -> LocalTime.NOON);
        try (var workbook = read(writer.write(values))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("1850-01-01", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("1899-01-01", sheet.getRow(2).getCell(0).getStringCellValue());
            assertEquals(1.0, sheet.getRow(3).getCell(0).getNumericCellValue());
            assertEquals(0.5, sheet.getRow(1).getCell(1).getNumericCellValue());
        }
    }

    @Test
    void preservesNumbersThatCannotBeWrittenExactlyWithTheInferredType() throws IOException {
        var rows = List.<Map<String, Object>>of(
                Map.of("Long", 9_007_199_254_740_993L, "Decimal", new BigDecimal("0.123456789012345678901"), "Mixed", 1),
                Map.of("Mixed", new BigDecimal("2.75")));
        var writer = ExcelWriter.forMap("Long", "Decimal", "Mixed")
                .writeErrorPolicy(ExcelWriteErrorPolicy.LENIENT);
        try (XSSFWorkbook workbook = read(writer.write(rows))) {
            var sheet = workbook.getSheetAt(0);
            assertEquals("9007199254740993", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("0.123456789012345678901", sheet.getRow(1).getCell(1).getStringCellValue());
            assertEquals(1.0, sheet.getRow(1).getCell(2).getNumericCellValue());
            assertEquals(2.75, sheet.getRow(2).getCell(2).getNumericCellValue());
        }
    }

    @Test
    void failFastRejectsLossyIntegerConversion() {
        var writer = ExcelWriter.<Number>create()
                .column("Value", value -> value, column -> column.type(ExcelDataType.INTEGER))
                .writeErrorPolicy(ExcelWriteErrorPolicy.FAIL_FAST);

        assertThrows(ExcelWriteException.class, () -> writer.write(List.of(1, 2.75)));
    }

    @Test
    void templateCellsPreserveLargeIntegersAsText() throws IOException {
        try (var workbook = new XSSFWorkbook()) {
            var cell = workbook.createSheet().createRow(0).createCell(0);
            ExcelTemplateWriter.setCellValue(cell, 9_007_199_254_740_993L);
            assertEquals("9007199254740993", cell.getStringCellValue());
        }
    }

    @Test
    void excelWriter_shouldInferCommonValueTypes() throws IOException {
        ExcelWriter<Map<String, Object>> writer = ExcelWriter.forMap(
                "Integer", "Long", "Double", "Decimal", "Date", "DateTime", "Time", "Boolean", "Text");

        ExcelHandler handler = writer.write(List.of(Map.of(
                "Integer", 12,
                "Long", 12L,
                "Double", 12.5d,
                "Decimal", new BigDecimal("12.50"),
                "Date", LocalDate.of(2026, 8, 6),
                "DateTime", LocalDateTime.of(2026, 8, 6, 13, 45, 30),
                "Time", LocalTime.of(13, 45, 30),
                "Boolean", true,
                "Text", "12"
        )));

        try (XSSFWorkbook workbook = read(handler)) {
            var row = workbook.getSheetAt(0).getRow(1);
            assertCellType(row.getCell(0), CellType.NUMERIC);
            assertCellType(row.getCell(1), CellType.NUMERIC);
            assertCellType(row.getCell(2), CellType.NUMERIC);
            assertCellType(row.getCell(3), CellType.NUMERIC);
            assertCellType(row.getCell(4), CellType.NUMERIC);
            assertCellType(row.getCell(5), CellType.NUMERIC);
            assertCellType(row.getCell(6), CellType.NUMERIC);
            assertCellType(row.getCell(7), CellType.STRING);
            assertEquals("Y", row.getCell(7).getStringCellValue());
            assertCellType(row.getCell(8), CellType.STRING);

            assertEquals("#,##0", row.getCell(0).getCellStyle().getDataFormatString());
            assertEquals("yyyy-mm-dd", row.getCell(4).getCellStyle().getDataFormatString());
            assertEquals("yyyy-mm-dd hh:mm:ss", row.getCell(5).getCellStyle().getDataFormatString());
            assertEquals("hh:mm:ss", row.getCell(6).getCellStyle().getDataFormatString());
        }
    }

    @Test
    void infer_shouldCoverRemainingSupportedAndFallbackTypes() {
        assertEquals(ExcelDataType.INTEGER, ExcelDataType.infer((byte) 1));
        assertEquals(ExcelDataType.INTEGER, ExcelDataType.infer((short) 1));
        assertEquals(ExcelDataType.FLOAT, ExcelDataType.infer(1.0f));
        assertEquals(ExcelDataType.STRING, ExcelDataType.infer(new StringBuilder("value")));
        assertEquals(ExcelDataType.HYPERLINK, ExcelDataType.infer(new ExcelHyperlink("https://example.com")));
        assertEquals(ExcelDataType.RICH_TEXT, ExcelDataType.infer(new ExcelRichText().text("value")));
        assertEquals(ExcelDataType.IMAGE, ExcelDataType.infer(ExcelImage.png(new byte[]{(byte) 137, 80, 78, 71})));
        assertNull(ExcelDataType.infer(null));
    }

    @Test
    void excelWriter_shouldInferAfterLeadingNull() throws IOException {
        ExcelWriter<Integer> writer = ExcelWriter.<Integer>create()
                .column("Value", value -> value == 0 ? null : value.longValue());

        ExcelHandler handler = writer.write(List.of(0, 7));
        try (XSSFWorkbook workbook = read(handler)) {
            var sheet = workbook.getSheetAt(0);
            assertCellType(sheet.getRow(1).getCell(0), CellType.STRING);
            assertCellType(sheet.getRow(2).getCell(0), CellType.NUMERIC);
            assertEquals(7L, (long) sheet.getRow(2).getCell(0).getNumericCellValue());
        }
    }

    @Test
    void excelWriter_shouldKeepStringWhenEveryValueIsNull() throws IOException {
        ExcelWriter<Object> writer = ExcelWriter.<Object>create().column("Value", value -> null);

        ExcelHandler handler = writer.write(List.of(new Object(), new Object()));
        try (XSSFWorkbook workbook = read(handler)) {
            var sheet = workbook.getSheetAt(0);
            assertCellType(sheet.getRow(1).getCell(0), CellType.STRING);
            assertEquals("", sheet.getRow(1).getCell(0).getStringCellValue());
            assertCellType(sheet.getRow(2).getCell(0), CellType.STRING);
        }
    }

    @Test
    void explicitType_shouldOverrideInferredValueType() throws IOException {
        ExcelWriter<String> writer = ExcelWriter.<String>create()
                .column("AsNumber", value -> 123, c -> c.type(ExcelDataType.INTEGER))
                .column("AsText", value -> 123, c -> c.type(ExcelDataType.STRING));

        ExcelHandler handler = writer.write(List.of("123"));
        try (XSSFWorkbook workbook = read(handler)) {
            var row = workbook.getSheetAt(0).getRow(1);
            assertCellType(row.getCell(0), CellType.NUMERIC);
            assertEquals(123, (int) row.getCell(0).getNumericCellValue());
            assertCellType(row.getCell(1), CellType.STRING);
            assertEquals("123", row.getCell(1).getStringCellValue());
        }
    }

    @Test
    void explicitFormat_shouldBePreservedForAutomaticallyInferredType() throws IOException {
        ExcelWriter<Void> writer = ExcelWriter.<Void>create()
                .column("Date", value -> LocalDate.of(2026, 8, 6), c -> c.format("dd/MM/yyyy"));

        ExcelHandler handler = writer.write(Stream.of((Void) null));
        try (XSSFWorkbook workbook = read(handler)) {
            Cell cell = workbook.getSheetAt(0).getRow(1).getCell(0);
            assertCellType(cell, CellType.NUMERIC);
            assertEquals("dd/MM/yyyy", cell.getCellStyle().getDataFormatString());
        }
    }

    @Test
    void inferredIntegerColumnPreservesLaterFractionAsNumeric() throws IOException {
        ExcelWriter<Object> writer = ExcelWriter.<Object>create()
                .column("Value", value -> value);

        ExcelHandler handler = writer.write(List.of(1, 2L, 3.5d));
        try (XSSFWorkbook workbook = read(handler)) {
            var sheet = workbook.getSheetAt(0);
            assertCellType(sheet.getRow(1).getCell(0), CellType.NUMERIC);
            assertCellType(sheet.getRow(2).getCell(0), CellType.NUMERIC);
            assertCellType(sheet.getRow(3).getCell(0), CellType.NUMERIC);
            assertEquals(3.5d, sheet.getRow(3).getCell(0).getNumericCellValue());
            assertEquals("#,##0", sheet.getRow(1).getCell(0).getCellStyle().getDataFormatString());
            assertEquals("General", sheet.getRow(3).getCell(0).getCellStyle().getDataFormatString());
        }
    }

    @Test
    void excelSheetWriter_shouldUseTheSameInferenceRules() throws IOException {
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<Object>sheet("Values")
                    .column("Date", value -> LocalDate.of(2026, 8, 6))
                    .column("Amount", value -> 42L)
                    .write(List.of(new Object()));

            ExcelHandler handler = workbook.finish();
            try (XSSFWorkbook result = read(handler)) {
                var row = result.getSheet("Values").getRow(1);
                assertCellType(row.getCell(0), CellType.NUMERIC);
                assertCellType(row.getCell(1), CellType.NUMERIC);
                assertEquals("yyyy-mm-dd", row.getCell(0).getCellStyle().getDataFormatString());
                assertEquals(42L, (long) row.getCell(1).getNumericCellValue());
            }
        }
    }

    private static void assertCellType(Cell cell, CellType expected) {
        assertEquals(expected, cell.getCellType(), "unexpected cell type for " + cell);
    }

    @Test
    void zonedAndLegacyDateTypesBecomeExcelDatesWhileInstantStaysText() throws IOException {
        var zone = java.time.ZoneId.systemDefault();
        var local = LocalDateTime.of(2025, 3, 4, 10, 30, 15);
        var epochMillis = local.atZone(zone).toInstant().toEpochMilli();
        var rows = java.util.Map.<String, Object>of(
                "zoned", local.atZone(java.time.ZoneId.of("Asia/Seoul")),
                "offset", local.atOffset(java.time.ZoneOffset.ofHours(-5)),
                "utilDate", new java.util.Date(epochMillis),
                "timestamp", new java.sql.Timestamp(epochMillis),
                "sqlDate", java.sql.Date.valueOf(local.toLocalDate()),
                "sqlTime", java.sql.Time.valueOf(local.toLocalTime()),
                "instant", java.time.Instant.parse("2025-03-04T01:30:15Z"));
        var writer = ExcelWriter.<Map<String, Object>>create()
                .column("zoned", r -> r.get("zoned")).column("offset", r -> r.get("offset"))
                .column("utilDate", r -> r.get("utilDate")).column("timestamp", r -> r.get("timestamp"))
                .column("sqlDate", r -> r.get("sqlDate")).column("sqlTime", r -> r.get("sqlTime"))
                .column("instant", r -> r.get("instant"));
        try (XSSFWorkbook workbook = read(writer.write(List.of(rows)))) {
            var row = workbook.getSheetAt(0).getRow(1);
            // Zoned values keep their own wall-clock time; legacy dates use the JVM default zone.
            assertEquals(local, row.getCell(0).getLocalDateTimeCellValue());
            assertEquals(local, row.getCell(1).getLocalDateTimeCellValue());
            assertEquals(local, row.getCell(2).getLocalDateTimeCellValue());
            assertEquals(local, row.getCell(3).getLocalDateTimeCellValue());
            assertEquals(local.toLocalDate().atStartOfDay(), row.getCell(4).getLocalDateTimeCellValue());
            assertEquals(local.toLocalTime().toSecondOfDay() / 86_400.0, row.getCell(5).getNumericCellValue(), 1e-9);
            // An Instant has no zone of its own, so it is not guessed.
            assertEquals(CellType.STRING, row.getCell(6).getCellType());
        }

        // The same types work in explicitly typed columns and template cells.
        var typed = ExcelWriter.<Object>create()
                .column("dt", v -> v, c -> c.type(ExcelDataType.DATETIME))
                .write(List.of(local.atZone(java.time.ZoneId.of("UTC"))));
        try (XSSFWorkbook workbook = read(typed)) {
            assertEquals(local, workbook.getSheetAt(0).getRow(1).getCell(0).getLocalDateTimeCellValue());
        }
    }

    private static XSSFWorkbook read(ExcelHandler handler) throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        handler.writeTo(output);
        assertTrue(output.size() > 0);
        return new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()));
    }
}
