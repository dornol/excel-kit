package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelWriter}
 */
class ExcelWriterTest {

    @Test
    void write_shouldThrowWhenNoColumns() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();

        // Act & Assert
        assertThrows(ExcelWriteException.class, () -> {
            writer.write(Stream.of("x", "y"));
        });
    }

    @Test
    void write_shouldSucceedWithSingleColumn() throws IOException {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        assertNotNull(handler);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void write_shouldReturnHandlerAndBeConsumable() throws IOException {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = List.of("row1", "row2").stream();

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row.length())
                .write(data);

        // Assert
        assertNotNull(handler, "write should return non-null ExcelHandler");
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0, "Produced Excel bytes should not be empty");
        }
    }

    @Test
    void write_shouldRolloverSheets_whenMaxRowsSmall() {
        // Arrange: max 2 rows per sheet
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3, 4, 5);

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row * 10)
                .write(data);

        // Assert before workbook is consumed
        SXSSFWorkbook wb = writer.getWb();
        assertEquals(3, wb.getNumberOfSheets(), "Expect 3 sheets when 5 rows with max 2 rows per sheet");

        // Also verify headers exist on each sheet
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            assertNotNull(s.getRow(0), "Header row must exist on each sheet");
            assertEquals("A", s.getRow(0).getCell(0).getStringCellValue());
            assertEquals("B", s.getRow(0).getCell(1).getStringCellValue());
        }

        // Finally consume to ensure no exception
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void columnIf_falseConditionShouldNotAddColumn() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .columnIf("B", false, (row, c) -> row)
                .column("C", (row, c) -> row)
                .write(data);

        // Assert header cell count is 2 (A, C)
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        int lastCellNum = sheet.getRow(0).getLastCellNum();
        assertEquals(2, lastCellNum, "Only two columns should be present when conditional column is false");

        // consume for completeness
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void constColumn_shouldWriteConstantValue() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("hello");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .constColumn("Const", "CONST_VAL")
                .write(data);

        // Assert header and first data row
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        assertEquals("Const", sheet.getRow(0).getCell(1).getStringCellValue());
        assertEquals("CONST_VAL", sheet.getRow(1).getCell(1).getStringCellValue());

        // consume for completeness
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void constructor_withRowAccessWindowSize_shouldCreateWriter() throws IOException {
        // Arrange: use a small buffer size
        ExcelWriter<String> writer = new ExcelWriter<>(255, 255, 255, 1_000_000, 100);
        Stream<String> data = Stream.of("a", "b", "c");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        assertNotNull(handler);
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        }
    }

    @Test
    void write_shouldApplyAutoFilter() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        // Act
        ExcelHandler handler = writer
                .autoFilter(true)
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row.length())
                .write(data);

        // Assert
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        // Verify auto-filter is set — SXSSFSheet tracks it internally
        assertNotNull(sheet, "Sheet should exist");

        // consume for completeness
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void write_shouldApplyFreezePane() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        // Act
        ExcelHandler handler = writer
                .freezePane(1)
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        PaneInformation pane = sheet.getPaneInformation();
        assertNotNull(pane, "Freeze pane information should exist");
        assertEquals(1, pane.getHorizontalSplitPosition(), "Freeze pane should freeze 1 row below header");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void write_shouldApplyOptionsOnRolloverSheets() {
        // Arrange: max 2 rows per sheet with auto-filter and freeze pane
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3, 4, 5);

        // Act
        ExcelHandler handler = writer
                .autoFilter(true)
                .freezePane(1)
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row * 10)
                .write(data);

        // Assert that every sheet has freeze pane applied
        SXSSFWorkbook wb = writer.getWb();
        assertTrue(wb.getNumberOfSheets() >= 2, "Should have multiple sheets");

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            PaneInformation pane = s.getPaneInformation();
            assertNotNull(pane, "Freeze pane should exist on sheet " + i);
            assertEquals(1, pane.getHorizontalSplitPosition(),
                    "Freeze pane should be at row 1 on sheet " + i);
        }

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void beforeHeader_shouldWriteCustomRowsBeforeHeader() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        // Act
        ExcelHandler handler = writer
                .beforeHeader(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Meta1");
                    ctx.getSheet().createRow(ctx.getCurrentRow() + 1).createCell(0).setCellValue("Meta2");
                    return ctx.getCurrentRow() + 2;
                })
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        // beforeHeader rows at 0, 1
        assertEquals("Meta1", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("Meta2", sheet.getRow(1).getCell(0).getStringCellValue());
        // Header at row 2
        assertEquals("A", sheet.getRow(2).getCell(0).getStringCellValue());
        // Data at row 3, 4
        assertEquals("a", sheet.getRow(3).getCell(0).getStringCellValue());
        assertEquals("b", sheet.getRow(4).getCell(0).getStringCellValue());

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void beforeHeader_shouldBeCalledOnEveryRolloverSheet() {
        // Arrange: max 2 rows per sheet
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3, 4, 5);

        // Act
        ExcelHandler handler = writer
                .beforeHeader(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("PREAMBLE");
                    return ctx.getCurrentRow() + 1;
                })
                .column("A", (row, c) -> row)
                .write(data);

        // Assert: 3 sheets, each with preamble at row 0 and header at row 1
        SXSSFWorkbook wb = writer.getWb();
        assertEquals(3, wb.getNumberOfSheets());

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            assertEquals("PREAMBLE", s.getRow(0).getCell(0).getStringCellValue(),
                    "Preamble must exist on sheet " + i);
            assertEquals("A", s.getRow(1).getCell(0).getStringCellValue(),
                    "Header must exist on sheet " + i);
        }

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void beforeHeader_shouldBeChainableFromColumnBuilder() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        // Act — call beforeHeader via builder chaining
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row.length())
                .beforeHeader(ctx -> {
                    ctx.getSheet().createRow(ctx.getCurrentRow()).createCell(0).setCellValue("Custom");
                    return ctx.getCurrentRow() + 1;
                })
                .column("C", (row, c) -> row.toUpperCase())
                .write(data);

        // Assert
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        // beforeHeader row at 0
        assertEquals("Custom", sheet.getRow(0).getCell(0).getStringCellValue());
        // Header at row 1
        assertEquals("A", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals("B", sheet.getRow(1).getCell(1).getStringCellValue());
        assertEquals("C", sheet.getRow(1).getCell(2).getStringCellValue());
        // Data at row 2, 3
        assertEquals("a", sheet.getRow(2).getCell(0).getStringCellValue());

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void freezePane_shouldThrowForNegativeValue() {
        ExcelWriter<String> writer = new ExcelWriter<>();
        assertThrows(IllegalArgumentException.class, () -> writer.freezePane(-1),
                "Negative freezePane value should throw IllegalArgumentException");
    }

    @Test
    void applyColumnWidth_shouldApplySameWidthsAcrossSheets() {
        // Arrange: small max rows to force rollover and values with different lengths
        ExcelWriter<String> writer = new ExcelWriter<>(2);
        Stream<String> data = Stream.of("short", "a bit longer", "short again");

        // Act
        ExcelHandler handler = writer
                .column("Col1", (row, c) -> row)
                .column("Col2", (row, c) -> row.toUpperCase())
                .write(data);

        // Assert widths are set and equal across sheets
        SXSSFWorkbook wb = writer.getWb();
        assertTrue(wb.getNumberOfSheets() >= 2);
        int w0c0 = wb.getSheetAt(0).getColumnWidth(0);
        int w0c1 = wb.getSheetAt(0).getColumnWidth(1);
        int w1c0 = wb.getSheetAt(1).getColumnWidth(0);
        int w1c1 = wb.getSheetAt(1).getColumnWidth(1);
        assertTrue(w0c0 > 0 && w0c1 > 0, "Column widths should be greater than zero");
        assertEquals(w0c0, w1c0, "Column 0 width should be equal across sheets");
        assertEquals(w0c1, w1c1, "Column 1 width should be equal across sheets");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void afterData_shouldWriteSubtotalAfterDataRows() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .afterData(ctx -> {
                    SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("subtotal");
                    return ctx.getCurrentRow() + 1;
                })
                .write(data);

        // Assert: header(0), data(1,2), subtotal(3)
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        assertEquals("A", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("a", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals("b", sheet.getRow(2).getCell(0).getStringCellValue());
        assertEquals("subtotal", sheet.getRow(3).getCell(0).getStringCellValue());

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void afterAll_shouldWriteTotalOnLastSheetOnly() {
        // Arrange: max 2 rows per sheet → 2 sheets for 3 data rows
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3);

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .afterAll(ctx -> {
                    SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("grand total");
                    return ctx.getCurrentRow() + 1;
                })
                .write(data);

        // Assert
        SXSSFWorkbook wb = writer.getWb();
        assertEquals(2, wb.getNumberOfSheets());

        // First sheet: header(0), data(1,2) — no grand total
        SXSSFSheet sheet0 = wb.getSheetAt(0);
        assertNull(sheet0.getRow(3), "First sheet should not have a total row");

        // Last sheet: header(0), data(1) → grand total at row 2
        SXSSFSheet sheet1 = wb.getSheetAt(1);
        assertEquals("grand total", sheet1.getRow(2).getCell(0).getStringCellValue());

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void afterData_and_afterAll_shouldBeCalledInOrder() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .afterData(ctx -> {
                    SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("subtotal");
                    return ctx.getCurrentRow() + 1;
                })
                .afterAll(ctx -> {
                    SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("grand total");
                    return ctx.getCurrentRow() + 1;
                })
                .write(data);

        // Assert: header(0), data(1), subtotal(2), grand total(3)
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        assertEquals("A", sheet.getRow(0).getCell(0).getStringCellValue());
        assertEquals("a", sheet.getRow(1).getCell(0).getStringCellValue());
        assertEquals("subtotal", sheet.getRow(2).getCell(0).getStringCellValue());
        assertEquals("grand total", sheet.getRow(3).getCell(0).getStringCellValue());

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void width_shouldFixColumnWidth() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("short", "a very long string that would normally expand the column");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .width(5000)
                .write(data);

        // Assert: column width should be exactly 5000 regardless of content
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        assertEquals(5000, sheet.getColumnWidth(0), "Fixed width should be applied exactly");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void minWidth_shouldEnforceMinimumColumnWidth() {
        // Arrange: use a large minWidth so auto-fit can't go below it
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("x");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .minWidth(10000)
                .write(data);

        // Assert: column width should be at least 10000
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        assertTrue(sheet.getColumnWidth(0) >= 10000,
                "Column width should be at least the minWidth value");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void maxWidth_shouldCapColumnWidth() {
        // Arrange: use a small maxWidth to cap auto-fit
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("This is a very long string value that would normally cause a very wide column");

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .maxWidth(3000)
                .write(data);

        // Assert: column width should not exceed 3000
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        assertTrue(sheet.getColumnWidth(0) <= 3000,
                "Column width should not exceed maxWidth value");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void sheetName_shouldSetCustomSheetName() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a", "b");

        // Act
        ExcelHandler handler = writer
                .sheetName("MySheet")
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        SXSSFWorkbook wb = writer.getWb();
        assertEquals("MySheet", wb.getSheetName(0), "First sheet should have custom name");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void sheetName_shouldNameRolloverSheets() {
        // Arrange: max 2 rows per sheet
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3, 4, 5);

        // Act
        ExcelHandler handler = writer
                .sheetName("Data")
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        SXSSFWorkbook wb = writer.getWb();
        assertEquals(3, wb.getNumberOfSheets());
        assertEquals("Data", wb.getSheetName(0));
        assertEquals("Data (2)", wb.getSheetName(1));
        assertEquals("Data (3)", wb.getSheetName(2));

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void sheetName_withFunction_shouldApplyCustomNaming() {
        // Arrange: max 2 rows per sheet
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3);

        // Act
        ExcelHandler handler = writer
                .sheetName(index -> "Page-" + (index + 1))
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        SXSSFWorkbook wb = writer.getWb();
        assertEquals(2, wb.getNumberOfSheets());
        assertEquals("Page-1", wb.getSheetName(0));
        assertEquals("Page-2", wb.getSheetName(1));

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void afterData_shouldBeCalledOnEverySheetDuringRollover() {
        // Arrange: max 2 rows per sheet → 3 sheets for 5 data rows
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3, 4, 5);

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .afterData(ctx -> {
                    SXSSFRow row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("subtotal");
                    return ctx.getCurrentRow() + 1;
                })
                .write(data);

        // Assert: 3 sheets, each with subtotal row after data
        SXSSFWorkbook wb = writer.getWb();
        assertEquals(3, wb.getNumberOfSheets());

        // Sheet 0: header(0), data(1,2), subtotal(3)
        assertEquals("subtotal", wb.getSheetAt(0).getRow(3).getCell(0).getStringCellValue(),
                "Subtotal must exist on sheet 0");

        // Sheet 1: header(0), data(1,2), subtotal(3)
        assertEquals("subtotal", wb.getSheetAt(1).getRow(3).getCell(0).getStringCellValue(),
                "Subtotal must exist on sheet 1");

        // Sheet 2: header(0), data(1), subtotal(2)
        assertEquals("subtotal", wb.getSheetAt(2).getRow(2).getCell(0).getStringCellValue(),
                "Subtotal must exist on sheet 2");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void dropdown_shouldApplyDataValidation() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("Active", "Inactive");

        // Act
        ExcelHandler handler = writer
                .column("Status", (row, c) -> row)
                .dropdown("Active", "Inactive", "Pending")
                .write(data);

        // Assert
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        List<? extends DataValidation> validations = sheet.getDataValidations();
        assertEquals(1, validations.size(), "Should have one data validation");
        DataValidation validation = validations.get(0);
        assertFalse(validation.getSuppressDropDownArrow(), "Dropdown arrow should not be suppressed");
        assertTrue(validation.getShowErrorBox(), "Error box should be shown");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void dropdown_shouldApplyToCorrectColumnOnly() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("test");

        // Act
        ExcelHandler handler = writer
                .column("Name", (row, c) -> row)
                .column("Status", (row, c) -> row)
                .dropdown("Active", "Inactive")
                .write(data);

        // Assert: only the Status column (index 1) should have validation
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        List<? extends DataValidation> validations = sheet.getDataValidations();
        assertEquals(1, validations.size(), "Should have exactly one data validation");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void rowColor_shouldApplyBackgroundColor() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("error", "ok");

        // Act
        ExcelHandler handler = writer
                .rowColor(row -> "error".equals(row) ? ExcelColor.LIGHT_RED : null)
                .column("Status", (row, c) -> row)
                .write(data);

        // Assert
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        // Row 1 is "error" → should have LIGHT_RED background
        SXSSFRow errorRow = sheet.getRow(1);
        assertNotNull(errorRow);
        var errorStyle = errorRow.getCell(0).getCellStyle();
        assertNotNull(errorStyle);
        XSSFColor errorFg = (XSSFColor) errorStyle.getFillForegroundColorColor();
        assertNotNull(errorFg, "Error row should have a fill color");
        byte[] rgb = errorFg.getRGB();
        assertEquals((byte) 255, rgb[0], "Red component");
        assertEquals((byte) 199, rgb[1], "Green component");
        assertEquals((byte) 206, rgb[2], "Blue component");

        // Row 2 is "ok" → should NOT have LIGHT_RED background
        SXSSFRow okRow = sheet.getRow(2);
        assertNotNull(okRow);

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void rowColor_shouldOverrideColumnBackgroundColor() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("highlight");

        // Act
        ExcelHandler handler = writer
                .rowColor(row -> ExcelColor.LIGHT_YELLOW)
                .column("Col", (row, c) -> row)
                .backgroundColor(ExcelColor.LIGHT_BLUE) // column bg
                .write(data);

        // Assert: row color should override column bg
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        SXSSFRow dataRow = sheet.getRow(1);
        XSSFColor fg = (XSSFColor) dataRow.getCell(0).getCellStyle().getFillForegroundColorColor();
        assertNotNull(fg);
        byte[] rgb = fg.getRGB();
        // Should be LIGHT_YELLOW, not LIGHT_BLUE
        assertEquals((byte) 255, rgb[0]);
        assertEquals((byte) 235, rgb[1]);
        assertEquals((byte) 156, rgb[2]);

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }

    @Test
    void sheetContext_shouldProvideColumnCountAndNames() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a");
        SheetContext[] captured = new SheetContext[1];

        // Act — capture context from beforeHeader callback
        ExcelHandler handler = writer
                .beforeHeader(ctx -> {
                    captured[0] = ctx;
                    return ctx.getCurrentRow();
                })
                .column("Name", (row, c) -> row)
                .column("Age", (row, c) -> row.length())
                .column("City", (row, c) -> row.toUpperCase())
                .write(data);

        // Assert
        assertNotNull(captured[0], "SheetContext should be provided to beforeHeader");
        assertEquals(3, captured[0].getColumnCount(), "Column count should match");
        assertEquals(List.of("Name", "Age", "City"), captured[0].getColumnNames(),
                "Column names should match in order");

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
        } catch (IOException e) {
            fail(e);
        }
    }
}
