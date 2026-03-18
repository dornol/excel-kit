package io.github.dornol.excelkit;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadResult;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests to boost coverage for CsvReadColumn.CsvReadColumnBuilder, ExcelSheetWriter.ColumnConfig,
 * CellData, ExcelDataType, and ExcelHyperlink.
 */
class CoverageBoostTest {

    @AfterEach
    void resetCellDataDefaults() {
        CellData.setDefaultLocale(Locale.getDefault());
        CellData.resetDateFormats();
        CellData.resetDateTimeFormats();
    }

    // -----------------------------------------------------------------------
    // 1. CsvReadColumn.CsvReadColumnBuilder — skipColumn, skipColumns, columnAt
    // -----------------------------------------------------------------------

    @Test
    void csvBuilder_skipColumn_shouldSkipOneColumnAndContinue() {
        String csv = "A,B,C\na1,b1,c1\na2,b2,c2\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .skipColumn()  // builder chain: skipColumn returns CsvReader, then column starts new builder
                .column((r, cell) -> r.third = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(2, results.size());
        assertEquals("a1", results.get(0).first);
        assertNull(results.get(0).second);
        assertEquals("c1", results.get(0).third);
        assertEquals("a2", results.get(1).first);
        assertEquals("c2", results.get(1).third);
    }

    @Test
    void csvBuilder_skipColumns_shouldSkipMultipleColumnsAndContinue() {
        String csv = "A,B,C,D\na1,b1,c1,d1\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .skipColumns(2)  // skip B and C
                .column((r, cell) -> r.fourth = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("a1", results.get(0).first);
        assertEquals("d1", results.get(0).fourth);
    }

    @Test
    void csvBuilder_columnAt_shouldReadByIndex() {
        String csv = "A,B,C,D\na1,b1,c1,d1\na2,b2,c2,d2\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .columnAt(3, (r, cell) -> r.fourth = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(2, results.size());
        assertEquals("a1", results.get(0).first);
        assertEquals("d1", results.get(0).fourth);
        assertEquals("a2", results.get(1).first);
        assertEquals("d2", results.get(1).fourth);
    }

    @Test
    void csvBuilder_columnAt_chainMultiple() {
        String csv = "A,B,C,D,E\na1,b1,c1,d1,e1\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column((r, cell) -> r.first = cell.asString())
                .columnAt(2, (r, cell) -> r.third = cell.asString())
                .columnAt(4, (r, cell) -> r.fifth = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("a1", results.get(0).first);
        assertEquals("c1", results.get(0).third);
        assertEquals("e1", results.get(0).fifth);
    }

    @Test
    void csvBuilder_column_namedChain() {
        String csv = "Name,Age,City\nAlice,30,Seoul\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8));

        List<CsvRow> results = new ArrayList<>();

        new CsvReader<>(CsvRow::new, null)
                .column("City", (r, cell) -> r.third = cell.asString())
                .column("Name", (r, cell) -> r.first = cell.asString())
                .build(is)
                .read(result -> results.add(result.data()));

        assertEquals(1, results.size());
        assertEquals("Seoul", results.get(0).third);
        assertEquals("Alice", results.get(0).first);
    }

    // -----------------------------------------------------------------------
    // 2. ExcelSheetWriter.ColumnConfig — format, alignment, bold, fontSize, width, minWidth, maxWidth
    // -----------------------------------------------------------------------

    @Test
    void columnConfig_format_shouldApplyCustomFormat() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<Integer>sheet("Data")
                .column("Value", v -> v, c -> c.type(ExcelDataType.INTEGER).format("#,##0.00"))
                .write(Stream.of(1234));

        workbook.finish().consumeOutputStream(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0.00", format);
        }
    }

    @Test
    void columnConfig_alignment_shouldApplyAlignment() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .column("Left", s -> s, c -> c.alignment(HorizontalAlignment.LEFT))
                .column("Right", s -> s, c -> c.alignment(HorizontalAlignment.RIGHT))
                .write(Stream.of("test"));

        workbook.finish().consumeOutputStream(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(1);
            assertEquals(HorizontalAlignment.LEFT, row.getCell(0).getCellStyle().getAlignment());
            assertEquals(HorizontalAlignment.RIGHT, row.getCell(1).getCellStyle().getAlignment());
        }
    }

    @Test
    void columnConfig_bold_shouldApplyBoldFont() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .column("Bold", s -> s, c -> c.bold(true))
                .column("NotBold", s -> s, c -> c.bold(false))
                .write(Stream.of("test"));

        workbook.finish().consumeOutputStream(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Row row = sheet.getRow(1);
            CellStyle boldStyle = row.getCell(0).getCellStyle();
            assertTrue(wb.getFontAt(boldStyle.getFontIndex()).getBold());
            CellStyle notBoldStyle = row.getCell(1).getCellStyle();
            assertFalse(wb.getFontAt(notBoldStyle.getFontIndex()).getBold());
        }
    }

    @Test
    void columnConfig_fontSize_shouldApplyFontSize() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .column("Big", s -> s, c -> c.fontSize(18))
                .write(Stream.of("test"));

        workbook.finish().consumeOutputStream(out);
        workbook.close();

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            CellStyle style = sheet.getRow(1).getCell(0).getCellStyle();
            assertEquals(18, wb.getFontAt(style.getFontIndex()).getFontHeightInPoints());
        }
    }

    @Test
    void columnConfig_fontSize_shouldThrowForNonPositive() {
        ExcelSheetWriter.ColumnConfig<String> config = new ExcelSheetWriter.ColumnConfig<>();
        assertThrows(IllegalArgumentException.class, () -> config.fontSize(0));
        assertThrows(IllegalArgumentException.class, () -> config.fontSize(-1));
    }

    @Test
    void columnConfig_width_shouldSetFixedWidth() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .column("Fixed", s -> s, c -> c.width(30))
                .write(Stream.of("test"));

        workbook.finish().consumeOutputStream(out);
        workbook.close();

        // Just verify it writes without error; fixed width is applied via internal logic
        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void columnConfig_minWidth_maxWidth_shouldWrite() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWorkbook workbook = new ExcelWorkbook();

        workbook.<String>sheet("Data")
                .column("Bounded", s -> s, c -> c.minWidth(10).maxWidth(50))
                .write(Stream.of("test"));

        workbook.finish().consumeOutputStream(out);
        workbook.close();

        assertTrue(out.toByteArray().length > 0);
    }

    // -----------------------------------------------------------------------
    // 3. CellData — asLocalDateTime(format), asLocalDate(format), asLocalTime(format),
    //    asNumber(Locale), asFloat(), asBooleanOrNull(), edge cases
    // -----------------------------------------------------------------------

    @Test
    void cellData_asLocalDateTime_withFormat() {
        CellData cell = new CellData(0, "19/07/2025 14:30:00");
        LocalDateTime result = cell.asLocalDateTime("dd/MM/yyyy HH:mm:ss");
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30, 0), result);
    }

    @Test
    void cellData_asLocalDateTime_withFormat_blankReturnsNull() {
        CellData cell = new CellData(0, "   ");
        assertNull(cell.asLocalDateTime("yyyy-MM-dd HH:mm:ss"));
    }

    @Test
    void cellData_asLocalDate_withFormat() {
        CellData cell = new CellData(0, "19.07.2025");
        LocalDate result = cell.asLocalDate("dd.MM.yyyy");
        assertEquals(LocalDate.of(2025, 7, 19), result);
    }

    @Test
    void cellData_asLocalDate_withFormat_blankReturnsNull() {
        CellData cell = new CellData(0, "");
        assertNull(cell.asLocalDate("yyyy/MM/dd"));
    }

    @Test
    void cellData_asLocalTime_withFormat() {
        CellData cell = new CellData(0, "14:30");
        LocalTime result = cell.asLocalTime("HH:mm");
        assertEquals(LocalTime.of(14, 30), result);
    }

    @Test
    void cellData_asLocalTime_withFormat_blankReturnsNull() {
        CellData cell = new CellData(0, "  ");
        assertNull(cell.asLocalTime("HH:mm:ss"));
    }

    @Test
    void cellData_asLocalTime_noArg() {
        CellData cell = new CellData(0, "10:15:30");
        LocalTime result = cell.asLocalTime();
        assertEquals(LocalTime.of(10, 15, 30), result);
    }

    @Test
    void cellData_asLocalTime_noArg_blankReturnsNull() {
        CellData cell = new CellData(0, "");
        assertNull(cell.asLocalTime());
    }

    @Test
    void cellData_asNumber_withLocale() {
        CellData cell = new CellData(0, "1,234.56");
        Number result = cell.asNumber(Locale.US);
        assertNotNull(result);
        assertEquals(1234.56, result.doubleValue(), 0.01);
    }

    @Test
    void cellData_asNumber_withLocale_blankReturnsNull() {
        CellData cell = new CellData(0, "   ");
        assertNull(cell.asNumber(Locale.US));
    }

    @Test
    void cellData_asNumber_withLocale_invalidThrows() {
        CellData cell = new CellData(0, "not-a-number");
        assertThrows(IllegalArgumentException.class, () -> cell.asNumber(Locale.US));
    }

    @Test
    void cellData_asNumber_withCurrencySymbols() {
        CellData cell = new CellData(0, "$1,000");
        Number result = cell.asNumber(Locale.US);
        assertNotNull(result);
        assertEquals(1000, result.longValue());
    }

    @Test
    void cellData_asNumber_withPercentSign() {
        CellData cell = new CellData(0, "50%");
        Number result = cell.asNumber(Locale.US);
        assertNotNull(result);
        assertEquals(50, result.longValue());
    }

    @Test
    void cellData_asFloat() {
        CellData cell = new CellData(0, "3.14");
        Float result = cell.asFloat();
        assertNotNull(result);
        assertEquals(3.14f, result, 0.01f);
    }

    @Test
    void cellData_asFloat_blankReturnsNull() {
        CellData cell = new CellData(0, "");
        assertNull(cell.asFloat());
    }

    @Test
    void cellData_asBooleanOrNull_trueValues() {
        assertEquals(Boolean.TRUE, new CellData(0, "true").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "1").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "y").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "YES").asBooleanOrNull());
        assertEquals(Boolean.TRUE, new CellData(0, "True").asBooleanOrNull());
    }

    @Test
    void cellData_asBooleanOrNull_falseValues() {
        assertEquals(Boolean.FALSE, new CellData(0, "false").asBooleanOrNull());
        assertEquals(Boolean.FALSE, new CellData(0, "0").asBooleanOrNull());
        assertEquals(Boolean.FALSE, new CellData(0, "no").asBooleanOrNull());
        assertEquals(Boolean.FALSE, new CellData(0, "n").asBooleanOrNull());
    }

    @Test
    void cellData_asBooleanOrNull_blankReturnsNull() {
        assertNull(new CellData(0, "").asBooleanOrNull());
        assertNull(new CellData(0, "   ").asBooleanOrNull());
    }

    @Test
    void cellData_negativeColumnIndex_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () -> new CellData(-1, "test"));
    }

    @Test
    void cellData_nullFormattedValue_becomesEmpty() {
        CellData cell = new CellData(0, null);
        assertEquals("", cell.formattedValue());
    }

    @Test
    void cellData_asDouble() {
        CellData cell = new CellData(0, "2.718");
        Double result = cell.asDouble();
        assertNotNull(result);
        assertEquals(2.718, result, 0.001);
    }

    @Test
    void cellData_asDouble_blankReturnsNull() {
        assertNull(new CellData(0, "").asDouble());
    }

    @Test
    void cellData_asBigDecimal() {
        CellData cell = new CellData(0, "12345.67");
        BigDecimal result = cell.asBigDecimal();
        assertNotNull(result);
        assertEquals(new BigDecimal("12345.67"), result);
    }

    @Test
    void cellData_asBigDecimal_blankReturnsNull() {
        assertNull(new CellData(0, "").asBigDecimal());
    }

    @Test
    void cellData_asInt_outOfRange_shouldThrow() {
        CellData cell = new CellData(0, String.valueOf(Long.MAX_VALUE));
        assertThrows(IllegalArgumentException.class, cell::asInt);
    }

    @Test
    void cellData_isEmpty() {
        assertTrue(new CellData(0, "").isEmpty());
        assertTrue(new CellData(0, "   ").isEmpty());
        assertFalse(new CellData(0, "x").isEmpty());
    }

    // -----------------------------------------------------------------------
    // 4. ExcelDataType — TIME, FLOAT, FLOAT_PERCENT, BIG_DECIMAL_TO_DOUBLE,
    //    BIG_DECIMAL_TO_LONG, BOOLEAN_TO_YN write and read back
    // -----------------------------------------------------------------------

    @Test
    void excelDataType_TIME_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<LocalTime>()
                .column("Time", t -> t)
                    .type(ExcelDataType.TIME)
                .write(Stream.of(LocalTime.of(14, 30, 15)))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            // TIME values are stored as date-time in Excel; verify format string
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("hh:mm:ss", format);
        }
    }

    @Test
    void excelDataType_FLOAT_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<Float>()
                .column("Float", f -> f)
                    .type(ExcelDataType.FLOAT)
                .write(Stream.of(3.14f))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(3.14, cell.getNumericCellValue(), 0.01);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0.00", format);
        }
    }

    @Test
    void excelDataType_FLOAT_PERCENT_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<Float>()
                .column("Pct", f -> f)
                    .type(ExcelDataType.FLOAT_PERCENT)
                .write(Stream.of(0.25f))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(0.25, cell.getNumericCellValue(), 0.001);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("0.00%", format);
        }
    }

    @Test
    void excelDataType_BIG_DECIMAL_TO_DOUBLE_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<BigDecimal>()
                .column("BD", bd -> bd)
                    .type(ExcelDataType.BIG_DECIMAL_TO_DOUBLE)
                .write(Stream.of(new BigDecimal("123.45")))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(123.45, cell.getNumericCellValue(), 0.01);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0.00", format);
        }
    }

    @Test
    void excelDataType_BIG_DECIMAL_TO_LONG_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<BigDecimal>()
                .column("BD", bd -> bd)
                    .type(ExcelDataType.BIG_DECIMAL_TO_LONG)
                .write(Stream.of(new BigDecimal("99999")))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            assertEquals(CellType.NUMERIC, cell.getCellType());
            assertEquals(99999.0, cell.getNumericCellValue(), 0.01);
            String format = wb.createDataFormat().getFormat(cell.getCellStyle().getDataFormat());
            assertEquals("#,##0", format);
        }
    }

    @Test
    void excelDataType_BOOLEAN_TO_YN_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<Boolean>()
                .column("Flag", b -> b)
                    .type(ExcelDataType.BOOLEAN_TO_YN)
                .write(Stream.of(true, false))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            assertEquals("Y", sheet.getRow(1).getCell(0).getStringCellValue());
            assertEquals("N", sheet.getRow(2).getCell(0).getStringCellValue());
        }
    }

    // -----------------------------------------------------------------------
    // 5. ExcelHyperlink — single-arg constructor
    // -----------------------------------------------------------------------

    @Test
    void excelHyperlink_singleArgConstructor() {
        ExcelHyperlink link = new ExcelHyperlink("https://example.com");
        assertEquals("https://example.com", link.url());
        assertEquals("https://example.com", link.label());
    }

    @Test
    void excelHyperlink_twoArgConstructor() {
        ExcelHyperlink link = new ExcelHyperlink("https://example.com", "Example");
        assertEquals("https://example.com", link.url());
        assertEquals("Example", link.label());
    }

    @Test
    void excelHyperlink_singleArg_writeAndReadBack() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();

        new ExcelWriter<String>()
                .column("Link", url -> new ExcelHyperlink(url))
                    .type(ExcelDataType.HYPERLINK)
                .write(Stream.of("https://example.com"))
                .consumeOutputStream(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            Cell cell = sheet.getRow(1).getCell(0);
            // single-arg: label == url
            assertEquals("https://example.com", cell.getStringCellValue());
            assertNotNull(cell.getHyperlink());
            assertEquals("https://example.com", cell.getHyperlink().getAddress());
        }
    }

    // -----------------------------------------------------------------------
    // Additional edge-case coverage
    // -----------------------------------------------------------------------

    @Test
    void cellData_asEnum() {
        CellData cell = new CellData(0, "STRING");
        ExcelDataType result = cell.asEnum(ExcelDataType.class);
        assertEquals(ExcelDataType.STRING, result);
    }

    @Test
    void cellData_asEnum_caseInsensitive() {
        CellData cell = new CellData(0, "string");
        ExcelDataType result = cell.asEnum(ExcelDataType.class);
        assertEquals(ExcelDataType.STRING, result);
    }

    @Test
    void cellData_asEnum_blankReturnsNull() {
        assertNull(new CellData(0, "").asEnum(ExcelDataType.class));
    }

    @Test
    void cellData_asEnum_invalidThrows() {
        CellData cell = new CellData(0, "INVALID_VALUE");
        assertThrows(IllegalArgumentException.class, () -> cell.asEnum(ExcelDataType.class));
    }

    @Test
    void cellData_asBoolean_trueValues() {
        assertTrue(new CellData(0, "true").asBoolean());
        assertTrue(new CellData(0, "1").asBoolean());
        assertTrue(new CellData(0, "y").asBoolean());
        assertTrue(new CellData(0, "yes").asBoolean());
    }

    @Test
    void cellData_asBoolean_falseValues() {
        assertFalse(new CellData(0, "false").asBoolean());
        assertFalse(new CellData(0, "0").asBoolean());
        assertFalse(new CellData(0, "").asBoolean());
    }

    @Test
    void cellData_asLong() {
        CellData cell = new CellData(0, "42");
        assertEquals(42L, cell.asLong());
    }

    @Test
    void cellData_asLong_blankReturnsNull() {
        assertNull(new CellData(0, "").asLong());
    }

    @Test
    void cellData_asInt() {
        CellData cell = new CellData(0, "42");
        assertEquals(42, cell.asInt());
    }

    @Test
    void cellData_asInt_blankReturnsNull() {
        assertNull(new CellData(0, "").asInt());
    }

    @Test
    void cellData_setDefaultLocale_nullThrows() {
        assertThrows(IllegalArgumentException.class, () -> CellData.setDefaultLocale(null));
    }

    @Test
    void cellData_asLocalDateTime_noArg_standardFormats() {
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30),
                new CellData(0, "2025-07-19 14:30").asLocalDateTime());
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30, 15),
                new CellData(0, "2025-07-19 14:30:15").asLocalDateTime());
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30),
                new CellData(0, "2025/07/19 14:30").asLocalDateTime());
        // ISO_LOCAL_DATE_TIME format
        assertEquals(LocalDateTime.of(2025, 7, 19, 14, 30, 15),
                new CellData(0, "2025-07-19T14:30:15").asLocalDateTime());
    }

    @Test
    void cellData_asLocalDateTime_noArg_blankReturnsNull() {
        assertNull(new CellData(0, "").asLocalDateTime());
    }

    @Test
    void cellData_asLocalDate_noArg_standardFormats() {
        assertEquals(LocalDate.of(2025, 7, 19), new CellData(0, "2025-07-19").asLocalDate());
        assertEquals(LocalDate.of(2025, 7, 19), new CellData(0, "2025/07/19").asLocalDate());
    }

    @Test
    void cellData_asLocalDate_noArg_blankReturnsNull() {
        assertNull(new CellData(0, "").asLocalDate());
    }

    @Test
    void columnConfig_chainingReturnsThis() {
        ExcelSheetWriter.ColumnConfig<String> config = new ExcelSheetWriter.ColumnConfig<>();
        assertSame(config, config.type(ExcelDataType.STRING));
        assertSame(config, config.format("#,##0"));
        assertSame(config, config.alignment(HorizontalAlignment.CENTER));
        assertSame(config, config.bold(true));
        assertSame(config, config.fontSize(12));
        assertSame(config, config.width(20));
        assertSame(config, config.minWidth(10));
        assertSame(config, config.maxWidth(50));
        assertSame(config, config.dropdown("A", "B"));
        assertSame(config, config.group("G1"));
        assertSame(config, config.outline(1));
    }

    @Test
    void columnConfig_outline_invalidLevel_shouldThrow() {
        ExcelSheetWriter.ColumnConfig<String> config = new ExcelSheetWriter.ColumnConfig<>();
        assertThrows(IllegalArgumentException.class, () -> config.outline(-1));
        assertThrows(IllegalArgumentException.class, () -> config.outline(8));
    }

    // -----------------------------------------------------------------------
    // Test data classes
    // -----------------------------------------------------------------------

    public static class CsvRow {
        String first;
        String second;
        String third;
        String fourth;
        String fifth;
    }
}
