package io.github.dornol.excelkit.excel;

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
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Tests value-based type inference for column writers. */
class ExcelAutomaticTypeInferenceTest {

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
    void inferredType_shouldBeStableAfterFirstNonNullValue() throws IOException {
        ExcelWriter<Object> writer = ExcelWriter.<Object>create()
                .column("Value", value -> value);

        ExcelHandler handler = writer.write(List.of(1, 2L, 3.5d));
        try (XSSFWorkbook workbook = read(handler)) {
            var sheet = workbook.getSheetAt(0);
            assertCellType(sheet.getRow(1).getCell(0), CellType.NUMERIC);
            assertCellType(sheet.getRow(2).getCell(0), CellType.NUMERIC);
            assertCellType(sheet.getRow(3).getCell(0), CellType.NUMERIC);
            assertEquals(3d, sheet.getRow(3).getCell(0).getNumericCellValue());
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

    private static XSSFWorkbook read(ExcelHandler handler) throws IOException {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        handler.writeTo(output);
        assertTrue(output.size() > 0);
        return new XSSFWorkbook(new ByteArrayInputStream(output.toByteArray()));
    }
}
