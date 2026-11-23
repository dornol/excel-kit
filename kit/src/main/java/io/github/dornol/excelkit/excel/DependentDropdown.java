package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.usermodel.SheetVisibility;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;

import java.util.ArrayList;
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/**
 * A dropdown whose options depend on the value of another column of the same row
 * (a parent value {@code A} offers {@code 1, 2, 3}; {@code B} offers {@code 4, 5, 6}).
 * <p>
 * The options are written to a hidden helper sheet (one column per parent value, the parent value in its first
 * row) and the column's list validation points at the column of the row's parent value.
 */
record DependentDropdown(String parentColumn, Map<String, List<String>> options) {

    private static final String SHEET_PREFIX = "_excelkit_dep_";
    /** Excel rejects a validation formula longer than this. */
    private static final int MAX_FORMULA_LENGTH = 255;
    /** The helper sheet already written for an option set, per workbook, so that rolled-over sheets share one. */
    private static final Map<SXSSFWorkbook, Map<Map<String, List<String>>, String>> HELPER_SHEETS =
            java.util.Collections.synchronizedMap(new java.util.WeakHashMap<>());
    private static final int MAX_COLUMNS = org.apache.poi.ss.SpreadsheetVersion.EXCEL2007.getMaxColumns();

    DependentDropdown {
        Objects.requireNonNull(parentColumn, "parentColumn cannot be null");
        Objects.requireNonNull(options, "options cannot be null");
        if (parentColumn.isBlank()) {
            throw new IllegalArgumentException("parentColumn must not be blank");
        }
        if (options.isEmpty()) {
            throw new IllegalArgumentException("A dependent dropdown needs at least one parent value");
        }
        if (options.size() > MAX_COLUMNS) {
            throw new IllegalArgumentException("A dependent dropdown has " + options.size()
                    + " parent values; a sheet has at most " + MAX_COLUMNS + " columns");
        }
        // an immutable, ordered, defensive copy
        var copy = new LinkedHashMap<String, List<String>>();
        for (var entry : options.entrySet()) {
            String key = Objects.requireNonNull(entry.getKey(), "a parent value cannot be null");
            if (key.isEmpty()) {
                throw new IllegalArgumentException("A parent value must not be empty");
            }
            List<String> children = Objects.requireNonNull(entry.getValue(),
                    "the options of parent value '" + key + "' cannot be null");
            if (children.isEmpty()) {
                throw new IllegalArgumentException("Parent value '" + key + "' needs at least one option");
            }
            if (children.size() > ExcelLimits.EXCEL_MAX_ROWS - 1) {
                throw new IllegalArgumentException("Parent value '" + key + "' has " + children.size()
                        + " options; a sheet has at most " + (ExcelLimits.EXCEL_MAX_ROWS - 1) + " rows below the header");
            }
            var childCopy = new ArrayList<String>(children.size());
            for (String child : children) {
                childCopy.add(Objects.requireNonNull(child, "an option of parent value '" + key + "' cannot be null"));
            }
            copy.put(key, List.copyOf(childCopy));
        }
        options = java.util.Collections.unmodifiableMap(copy);
    }

    /**
     * Applies this dropdown to the column {@code col} of {@code sheet}.
     *
     * @param parentCol the 0-based index of the parent column
     * @param headerRow the 0-based index of the header row
     */
    void apply(SXSSFSheet sheet, int col, int parentCol, int headerRow) {
        SXSSFWorkbook wb = sheet.getWorkbook();
        String helperName = HELPER_SHEETS.computeIfAbsent(wb, k -> new java.util.HashMap<>())
                .computeIfAbsent(options, k -> writeHelperSheet(wb));

        int keys = options.size();
        int maxChildren = options.values().stream().mapToInt(List::size).max().orElse(1);
        String ref = "'" + helperName + "'!";
        String lastKeyCol = SheetContext.columnLetter(keys - 1);
        // The row of the first data cell, 1-based; relative so that every row reads the parent value of its own row.
        String parentCell = "$" + SheetContext.columnLetter(parentCol) + (headerRow + 2);
        String offset = "MATCH(" + parentCell + "," + ref + "$A$1:$" + lastKeyCol + "$1,0)-1";
        String formula = "OFFSET(" + ref + "$A$2,0," + offset + ",COUNTA(OFFSET(" + ref + "$A$2,0," + offset + ","
                + maxChildren + ",1)),1)";
        if (formula.length() > MAX_FORMULA_LENGTH) {
            throw new IllegalArgumentException("The dependent dropdown formula is " + formula.length()
                    + " characters, over Excel's limit of " + MAX_FORMULA_LENGTH);
        }

        DataValidationHelper helper = sheet.getDataValidationHelper();
        DataValidation validation = helper.createValidation(helper.createFormulaListConstraint(formula),
                new CellRangeAddressList(headerRow + 1, ExcelLimits.EXCEL_MAX_ROWS, col, col));
        validation.setSuppressDropDownArrow(false); // keep the in-cell arrow
        validation.setShowErrorBox(true);
        sheet.addValidationData(validation);
    }

    /** Writes the parent values and their options to a new hidden sheet and returns its name. */
    private String writeHelperSheet(SXSSFWorkbook wb) {
        int n = 1;
        while (wb.getSheet(SHEET_PREFIX + n) != null) n++;
        String name = SHEET_PREFIX + n;
        SXSSFSheet helperSheet = wb.createSheet(name);
        int maxChildren = options.values().stream().mapToInt(List::size).max().orElse(1);
        var header = helperSheet.createRow(0);
        var rows = new ArrayList<org.apache.poi.ss.usermodel.Row>(maxChildren);
        for (int r = 0; r < maxChildren; r++) rows.add(helperSheet.createRow(r + 1));
        int c = 0;
        for (var entry : options.entrySet()) {
            header.createCell(c).setCellValue(entry.getKey());
            List<String> children = entry.getValue();
            for (int r = 0; r < children.size(); r++) {
                rows.get(r).createCell(c).setCellValue(children.get(r));
            }
            c++;
        }
        wb.setSheetVisibility(wb.getSheetIndex(helperSheet), SheetVisibility.HIDDEN);
        return name;
    }
}
