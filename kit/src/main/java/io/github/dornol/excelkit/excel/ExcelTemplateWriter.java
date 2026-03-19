package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.jspecify.annotations.Nullable;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.io.InputStream;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.util.HashMap;
import java.util.Map;

/**
 * Fills data into an existing Excel template (.xlsx) while preserving formatting,
 * images, charts, and merged regions.
 * <p>
 * Supports two write modes that can be combined:
 * <ul>
 *   <li><b>Cell-level writes</b> — fill individual cells by reference (e.g., "B3")</li>
 *   <li><b>List streaming</b> — write tabular data starting from a given row</li>
 * </ul>
 *
 * <p><b>Important:</b> All writes must proceed in top-to-bottom row order.
 * SXSSFWorkbook flushes rows from memory, so previously written rows cannot be revisited.
 *
 * <pre>{@code
 * try (ExcelTemplateWriter writer = new ExcelTemplateWriter(templateStream)) {
 *     writer.cell("B3", clientName)
 *           .cell("B4", LocalDate.now());
 *
 *     writer.<Item>list(5)
 *           .column("A", Item::getName)
 *           .column("B", Item::getQty)
 *           .column("C", Item::getAmount)
 *           .afterData(ctx -> {
 *               ctx.getSheet().createRow(ctx.getCurrentRow())
 *                  .createCell(0).setCellValue("Total: " + total);
 *               return ctx.getCurrentRow() + 1;
 *           })
 *           .write(itemStream);
 *
 *     writer.finish().consumeOutputStream(outputStream);
 * }
 * }</pre>
 *
 * @author dhkim
 * @since 0.8.2
 */
public class ExcelTemplateWriter implements AutoCloseable {
    private static final Logger log = LoggerFactory.getLogger(ExcelTemplateWriter.class);
    private static final int DEFAULT_ROW_ACCESS_WINDOW_SIZE = 1000;

    private final XSSFWorkbook templateWb;
    private final SXSSFWorkbook wb;
    private final Map<String, CellStyle> cellStyleCache = new HashMap<>();
    private boolean finished = false;
    private int activeSheetIndex = 0;
    private final Map<Integer, Integer> lastWrittenRowBySheet = new HashMap<>();
    private final Map<Integer, Integer> templateLastRowBySheet = new HashMap<>();

    /**
     * Opens a template from the given input stream.
     *
     * @param templateStream the .xlsx template file stream
     * @throws IOException if the template cannot be read
     */
    public ExcelTemplateWriter(InputStream templateStream) throws IOException {
        this(templateStream, DEFAULT_ROW_ACCESS_WINDOW_SIZE);
    }

    /**
     * Opens a template from the given input stream with a custom row access window size.
     *
     * @param templateStream      the .xlsx template file stream
     * @param rowAccessWindowSize number of rows kept in memory by SXSSFWorkbook
     * @throws IOException if the template cannot be read
     */
    public ExcelTemplateWriter(InputStream templateStream, int rowAccessWindowSize) throws IOException {
        this.templateWb = new XSSFWorkbook(templateStream);
        this.wb = new SXSSFWorkbook(templateWb, rowAccessWindowSize);
        // Record last existing row per sheet so we know which rows are in the template
        for (int i = 0; i < templateWb.getNumberOfSheets(); i++) {
            int lastRow = templateWb.getSheetAt(i).getLastRowNum();
            if (templateWb.getSheetAt(i).getPhysicalNumberOfRows() > 0) {
                templateLastRowBySheet.put(i, lastRow);
            }
        }
    }

    /**
     * Selects the active sheet by index for subsequent {@link #cell} calls.
     *
     * @param index 0-based sheet index
     * @return this writer for chaining
     */
    public ExcelTemplateWriter sheet(int index) {
        if (index < 0 || index >= wb.getNumberOfSheets()) {
            throw new ExcelWriteException("Sheet index out of range: " + index
                    + " (workbook has " + wb.getNumberOfSheets() + " sheets)");
        }
        this.activeSheetIndex = index;
        return this;
    }

    /**
     * Selects the active sheet by name for subsequent {@link #cell} calls.
     *
     * @param name the sheet name
     * @return this writer for chaining
     */
    public ExcelTemplateWriter sheet(String name) {
        int index = wb.getSheetIndex(name);
        if (index < 0) {
            throw new ExcelWriteException("Sheet not found: " + name);
        }
        this.activeSheetIndex = index;
        return this;
    }

    /**
     * Writes a value to the specified cell on the active sheet.
     * <p>
     * The value type is auto-detected: String, Number, Boolean, LocalDate, LocalDateTime,
     * LocalTime, or null (blank cell).
     *
     * @param cellRef Excel-notation cell reference (e.g., "B3", "AA10")
     * @param value   the value to write
     * @return this writer for chaining
     */
    public ExcelTemplateWriter cell(String cellRef, @Nullable Object value) {
        CellReference ref = new CellReference(cellRef);
        return cell(ref.getRow(), ref.getCol(), value, null);
    }

    /**
     * Writes a value with a custom style to the specified cell.
     *
     * @param cellRef Excel-notation cell reference
     * @param value   the value to write
     * @param style   the cell style to apply
     * @return this writer for chaining
     */
    public ExcelTemplateWriter cell(String cellRef, @Nullable Object value, CellStyle style) {
        CellReference ref = new CellReference(cellRef);
        return cell(ref.getRow(), ref.getCol(), value, style);
    }

    /**
     * Writes a value to the specified cell by row and column index.
     *
     * @param row   0-based row index
     * @param col   0-based column index
     * @param value the value to write
     * @return this writer for chaining
     */
    public ExcelTemplateWriter cell(int row, int col, @Nullable Object value) {
        return cell(row, col, value, null);
    }

    /**
     * Writes a value with a custom style to the specified cell by row and column index.
     *
     * @param row   0-based row index
     * @param col   0-based column index
     * @param value the value to write
     * @param style the cell style to apply (null to keep existing)
     * @return this writer for chaining
     */
    public ExcelTemplateWriter cell(int row, int col, @Nullable Object value, @Nullable CellStyle style) {
        checkNotFinished();
        enforceRowOrder(activeSheetIndex, row);

        int templateLastRow = templateLastRowBySheet.getOrDefault(activeSheetIndex, -1);
        if (row <= templateLastRow) {
            // Row exists in template — write via underlying XSSFSheet to avoid SXSSFWorkbook flush conflict
            org.apache.poi.xssf.usermodel.XSSFSheet xssfSheet = templateWb.getSheetAt(activeSheetIndex);
            Row sheetRow = xssfSheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = xssfSheet.createRow(row);
            }
            Cell cell = sheetRow.getCell(col);
            if (cell == null) {
                cell = sheetRow.createCell(col);
            }
            setCellValue(cell, value);
            if (style != null) {
                cell.setCellStyle(style);
            }
        } else {
            // New row beyond template — write via SXSSFWorkbook
            SXSSFSheet sheet = wb.getSheetAt(activeSheetIndex);
            Row sheetRow = sheet.getRow(row);
            if (sheetRow == null) {
                sheetRow = sheet.createRow(row);
            }
            Cell cell = sheetRow.getCell(col);
            if (cell == null) {
                cell = sheetRow.createCell(col);
            }
            setCellValue(cell, value);
            if (style != null) {
                cell.setCellStyle(style);
            }
        }
        return this;
    }

    /**
     * Creates a list writer for streaming tabular data starting at the given row on the active sheet.
     * <p>
     * The template's existing column headers (if any) are preserved.
     * Data rows are written starting from {@code startRow}.
     *
     * @param startRow 0-based row index where data writing begins
     * @param <T>      the row data type
     * @return a {@link TemplateListWriter} for configuring columns and writing data
     */
    public <T> TemplateListWriter<T> list(int startRow) {
        return list(activeSheetIndex, startRow);
    }

    /**
     * Creates a list writer for streaming tabular data into a specific sheet.
     *
     * @param sheetIndex 0-based sheet index
     * @param startRow   0-based row index where data writing begins
     * @param <T>        the row data type
     * @return a {@link TemplateListWriter} for configuring columns and writing data
     */
    public <T> TemplateListWriter<T> list(int sheetIndex, int startRow) {
        checkNotFinished();
        enforceRowOrder(sheetIndex, startRow);
        SXSSFSheet sheet = wb.getSheetAt(sheetIndex);
        return new TemplateListWriter<>(this, wb, sheet, startRow, cellStyleCache, sheetIndex);
    }

    /**
     * Finishes the template and returns an {@link ExcelHandler} for output.
     * <p>
     * After calling this method, no more writes are allowed.
     *
     * @return ExcelHandler wrapping the workbook
     */
    public ExcelHandler finish() {
        checkNotFinished();
        finished = true;
        return new ExcelHandler(wb);
    }

    /**
     * Closes the underlying workbook if it has not been finished.
     */
    @Override
    public void close() {
        if (!finished) {
            try {
                wb.close();
            } catch (Exception e) {
                log.warn("Failed to close template workbook", e);
            }
        }
    }

    void updateLastWrittenRow(int sheetIndex, int row) {
        lastWrittenRowBySheet.merge(sheetIndex, row, Math::max);
    }

    private void enforceRowOrder(int sheetIndex, int row) {
        int lastRow = lastWrittenRowBySheet.getOrDefault(sheetIndex, -1);
        if (row < lastRow) {
            throw new ExcelWriteException(
                    "Cells must be written in top-down row order. "
                            + "Attempted row " + row + " but last written row was " + lastRow
                            + " on sheet " + sheetIndex);
        }
        lastWrittenRowBySheet.put(sheetIndex, row);
    }

    private void checkNotFinished() {
        if (finished) {
            throw new ExcelWriteException("Template writer is already finished");
        }
    }

    static void setCellValue(Cell cell, @Nullable Object value) {
        if (value == null) {
            cell.setBlank();
        } else if (value instanceof String s) {
            cell.setCellValue(s);
        } else if (value instanceof Number n) {
            cell.setCellValue(n.doubleValue());
        } else if (value instanceof Boolean b) {
            cell.setCellValue(b);
        } else if (value instanceof LocalDateTime ldt) {
            cell.setCellValue(ldt);
        } else if (value instanceof LocalDate ld) {
            cell.setCellValue(ld);
        } else if (value instanceof LocalTime lt) {
            cell.setCellValue(lt.atDate(LocalDate.EPOCH));
        } else {
            cell.setCellValue(String.valueOf(value));
        }
    }
}
