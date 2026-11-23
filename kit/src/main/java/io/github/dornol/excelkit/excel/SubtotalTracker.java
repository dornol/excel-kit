package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;

import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.Function;

/**
 * Writes a subtotal row below every run of rows that share a key. The rows of a group are contiguous, so the
 * range of its {@code SUBTOTAL} is known when the group ends and nothing is buffered. One tracker lives for one
 * write; a group is closed at the end of a sheet, so a rolled-over sheet starts a new group.
 */
final class SubtotalTracker<T> {
    private final Function<T, @Nullable Object> key;
    private final ExcelSubtotal config;
    /** The style of the subtotal rows; locked, so that a column a person may edit does not let them overwrite a formula. */
    private final RowStyleConfig style;
    private final List<ExcelColumn<T>> columns;
    private final int labelIndex;
    private final ExcelSummary.@Nullable Op[] opOfColumn;

    private boolean open;
    private int startRow;
    private @Nullable Object groupKey;

    private SubtotalTracker(Function<T, @Nullable Object> key, ExcelSubtotal config, List<ExcelColumn<T>> columns) {
        this.key = key;
        this.config = config;
        this.style = config.style().lockedCopy();
        this.columns = columns;
        List<String> names = columns.stream().map(ExcelColumn::getName).toList();
        String labelColumn = config.labelColumn();
        this.labelIndex = labelColumn == null ? 0 : names.indexOf(labelColumn);
        if (labelIndex < 0) throw new ExcelWriteException("Subtotal label column not found: '" + labelColumn
                + "'. Columns: " + names);
        this.opOfColumn = new ExcelSummary.Op[columns.size()];
        for (var entry : config.entries().entrySet()) {
            for (String name : entry.getValue()) {
                int index = names.indexOf(name);
                if (index < 0) throw new ExcelWriteException("Subtotal column not found: '" + name
                        + "'. Columns: " + names);
                if (index == labelIndex) throw new ExcelWriteException("Subtotal column '" + name
                        + "' also holds the label of the subtotal row");
                if (opOfColumn[index] != null) throw new ExcelWriteException("Subtotal column '" + name
                        + "' is given more than one function");
                opOfColumn[index] = entry.getKey();
            }
        }
    }

    /** The tracker, or {@code null} when the sheet has no subtotals. */
    static <T> @Nullable SubtotalTracker<T> of(SheetConfig<T> cfg, List<ExcelColumn<T>> columns) {
        if (cfg.subtotalKey == null || cfg.subtotalConfig == null) return null;
        return new SubtotalTracker<>(cfg.subtotalKey, cfg.subtotalConfig, columns);
    }

    /** Rejects what a subtotal row would corrupt: tables, charts and named ranges span the rows as one data block. */
    static <T> void validate(SheetConfig<T> cfg) {
        if (cfg.subtotalKey == null) return;
        if (cfg.tableOptions != null) throw new ExcelWriteException(
                "groupSubtotal(...) cannot be combined with table(...): the subtotal rows would become rows of the table");
        if (cfg.chartConfig != null) throw new ExcelWriteException(
                "groupSubtotal(...) cannot be combined with chart(...): the chart would plot the subtotal rows");
        if (cfg.namedRanges != null && !cfg.namedRanges.isEmpty()) throw new ExcelWriteException(
                "groupSubtotal(...) cannot be combined with namedRange(...): the range would include the subtotal rows");
    }

    /**
     * Called before {@code row} is written: ends the open group when the key changes, then starts the group
     * of {@code row} when none is open.
     */
    void beforeRow(SXSSFSheet sheet, Cursor cursor, T row, SheetConfig<T> cfg, @Nullable MergeRunTracker<T> merges,
                   Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> styleCache, SXSSFWorkbook wb) {
        Object current = key.apply(row);
        if (open && !Objects.equals(groupKey, current)) close(sheet, cursor, cfg, merges, styleCache, wb);
        if (!open) {
            open = true;
            startRow = cursor.getRowOfSheet();
            groupKey = current;
        }
    }

    /** Ends the open group, if any, with its subtotal row: at a key change, at the end of a sheet and of the data. */
    void close(SXSSFSheet sheet, Cursor cursor, SheetConfig<T> cfg, @Nullable MergeRunTracker<T> merges,
               Map<ExcelWriteSupport.DerivedStyleKey, CellStyle> styleCache, SXSSFWorkbook wb) {
        if (!open) return;
        open = false;
        int endRow = cursor.getRowOfSheet() - 1;
        // A run of equal values must not carry on into the subtotal row.
        if (merges != null) merges.closeRuns(endRow);
        if (cursor.getRowOfSheet() >= ExcelLimits.SHEET_ROW_CAPACITY) {
            throw new ExcelWriteException("Sheet '" + sheet.getSheetName() + "' is full: no room for the subtotal row. "
                    + "Lower maxRows(...) so the writer rolls over to a new sheet earlier.");
        }
        SXSSFRow target = sheet.createRow(cursor.getRowOfSheet());
        target.setHeightInPoints(cfg.rowHeightInPoints);
        cursor.plusRow();
        for (int i = 0; i < columns.size(); i++) {
            SXSSFCell cell = target.createCell(i);
            ExcelSummary.Op op = opOfColumn[i];
            if (i == labelIndex) {
                String text = config.labelText().apply(groupKey);
                ExcelCellText.write(cell, text);
                // a label wider than the data of its column widens it, within the rows the width is sampled from
                if (cfg.autoWidthSampleRows > 0 && cursor.getCurrentTotal() <= cfg.autoWidthSampleRows) {
                    columns.get(i).fitColumnWidthByValue(text);
                }
            } else if (op != null) {
                String letter = SheetContext.columnLetter(i);
                ExcelDataType.setUncalculatedFormula(cell, "SUBTOTAL(" + ExcelSubtotal.functionNumber(op) + ","
                        + letter + (startRow + 1) + ":" + letter + (endRow + 1) + ")");
            } else {
                cell.setBlank();
            }
            cell.setCellStyle(ExcelWriteSupport.resolveRowStyle(columns.get(i).getStyle(), style.backgroundColor,
                    style, styleCache, wb, cfg.maxCellStyles));
        }
    }
}
