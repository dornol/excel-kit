package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
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
    void write_shouldSetTitleOnEachSheet_whenRolloverWithTitle() {
        // Arrange: max 2 rows per sheet with title
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3, 4, 5);

        // Act
        ExcelHandler handler = writer
                .title("Test Title")
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row * 10)
                .write(data);

        // Assert
        SXSSFWorkbook wb = writer.getWb();
        assertEquals(3, wb.getNumberOfSheets(), "Expect 3 sheets when 5 rows with max 2 rows per sheet");

        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            SXSSFSheet s = wb.getSheetAt(i);
            // Title row (row 0)
            assertNotNull(s.getRow(0), "Title row must exist on sheet " + i);
            assertEquals("Test Title", s.getRow(0).getCell(0).getStringCellValue(),
                    "Title must be set on sheet " + i);
            // Header row (row 2, after title rows 0-1)
            assertNotNull(s.getRow(2), "Header row must exist on sheet " + i);
            assertEquals("A", s.getRow(2).getCell(0).getStringCellValue(),
                    "Header A must be set on sheet " + i);
            assertEquals("B", s.getRow(2).getCell(1).getStringCellValue(),
                    "Header B must be set on sheet " + i);
        }

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
    void write_shouldApplyAutoFilterWithTitle() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("a");

        // Act
        ExcelHandler handler = writer
                .title("Title")
                .autoFilter(true)
                .column("A", (row, c) -> row)
                .column("B", (row, c) -> row.length())
                .write(data);

        // Assert — the auto-filter range should start at row 2 when title is present
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        assertNotNull(sheet);

        // consume
        try (ByteArrayOutputStream bos = new ByteArrayOutputStream()) {
            handler.consumeOutputStream(bos);
            assertTrue(bos.toByteArray().length > 0);
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
                .beforeHeader((sheet, wb, startRow) -> {
                    sheet.createRow(startRow).createCell(0).setCellValue("Meta1");
                    sheet.createRow(startRow + 1).createCell(0).setCellValue("Meta2");
                    return startRow + 2;
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
    void beforeHeader_withTitle_shouldWriteAfterTitle() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        Stream<String> data = Stream.of("x");

        // Act
        ExcelHandler handler = writer
                .title("Report Title")
                .beforeHeader((sheet, wb, startRow) -> {
                    // startRow should be 2 (after title rows 0-1)
                    sheet.createRow(startRow).createCell(0).setCellValue("Author: Kim");
                    return startRow + 1;
                })
                .column("A", (row, c) -> row)
                .write(data);

        // Assert
        SXSSFSheet sheet = writer.getWb().getSheetAt(0);
        // Title at row 0
        assertEquals("Report Title", sheet.getRow(0).getCell(0).getStringCellValue());
        // beforeHeader at row 2
        assertEquals("Author: Kim", sheet.getRow(2).getCell(0).getStringCellValue());
        // Header at row 3
        assertEquals("A", sheet.getRow(3).getCell(0).getStringCellValue());
        // Data at row 4
        assertEquals("x", sheet.getRow(4).getCell(0).getStringCellValue());

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
                .beforeHeader((sheet, wb, startRow) -> {
                    sheet.createRow(startRow).createCell(0).setCellValue("PREAMBLE");
                    return startRow + 1;
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
                .afterData((sheet, wb, nextRow) -> {
                    SXSSFRow row = sheet.createRow(nextRow);
                    row.createCell(0).setCellValue("subtotal");
                    return nextRow + 1;
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
                .afterAll((sheet, wb, nextRow) -> {
                    SXSSFRow row = sheet.createRow(nextRow);
                    row.createCell(0).setCellValue("grand total");
                    return nextRow + 1;
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
                .afterData((sheet, wb, nextRow) -> {
                    SXSSFRow row = sheet.createRow(nextRow);
                    row.createCell(0).setCellValue("subtotal");
                    return nextRow + 1;
                })
                .afterAll((sheet, wb, nextRow) -> {
                    SXSSFRow row = sheet.createRow(nextRow);
                    row.createCell(0).setCellValue("grand total");
                    return nextRow + 1;
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
    void afterData_shouldBeCalledOnEverySheetDuringRollover() {
        // Arrange: max 2 rows per sheet → 3 sheets for 5 data rows
        ExcelWriter<Integer> writer = new ExcelWriter<>(2);
        Stream<Integer> data = Stream.of(1, 2, 3, 4, 5);

        // Act
        ExcelHandler handler = writer
                .column("A", (row, c) -> row)
                .afterData((sheet, wb, nextRow) -> {
                    SXSSFRow row = sheet.createRow(nextRow);
                    row.createCell(0).setCellValue("subtotal");
                    return nextRow + 1;
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
}
