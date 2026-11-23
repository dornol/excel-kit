package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelCellComment;
import io.github.dornol.excelkit.core.Cursor;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.jspecify.annotations.Nullable;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Objects;

/** Writes the header rows of a sheet: the column names, the group rows above them and their merges. */
final class ExcelHeaderWriter {
    private ExcelHeaderWriter() { }

    /** A text format that shows {@code *} after the text of the cell. */
    private static final String REQUIRED_MARK_FORMAT = "@\" *\"";

    /**
     * Writes column headers, with 0..N optional group header rows if any column has groups.
     * {@code workbook} and {@code styleCache} are needed only for columns with their own header colors.
     *
     * @param tableHeader whether the column names row becomes the header of a table. Excel requires a table's header
     *                    cells to hold the column names, so a column without a group (whose name would sit in the
     *                    first header row, merged down) gets its name in the column names row instead, unmerged
     */
    static <T> void write(SXSSFSheet sheet, Cursor cursor,
                                        List<ExcelColumn<T>> columns, CellStyle headerStyle,
                                        @Nullable SXSSFWorkbook wb, @Nullable Map<String, CellStyle> headerStyleCache,
                                        @Nullable Map<List<String>, ExcelCellComment> groupComments,
                                        float headerRowHeight, boolean tableHeader) {
        int maxDepth = 0;
        for (ExcelColumn<T> c : columns) {
            int d = c.getGroupNames().length;
            if (d > maxDepth) maxDepth = d;
        }
        if (maxDepth == 0) {
            writeSingleHeaderRow(sheet, cursor, columns, headerStyle, wb, headerStyleCache, headerRowHeight);
        } else {
            writeGroupAndColumnHeaders(sheet, cursor, columns, headerStyle, wb, headerStyleCache,
                    maxDepth, groupComments, headerRowHeight, tableHeader);
        }
    }

    private static <T> CellStyle resolveHeaderStyle(ExcelColumn<T> col, CellStyle baseStyle,
                                                     @Nullable SXSSFWorkbook wb,
                                                     @Nullable Map<String, CellStyle> cache) {
        int[] fontColor = col.getHeaderFontColor();
        int[] bgColor = col.getHeaderBackgroundColor();
        boolean required = col.isRequiredMark();
        if ((fontColor == null && bgColor == null && !required) || wb == null || cache == null) {
            return baseStyle;
        }
        String key = "hdr_" + baseStyle.getIndex()
                + "_f" + (fontColor == null ? "-" : fontColor[0] + "_" + fontColor[1] + "_" + fontColor[2])
                + "_b" + (bgColor == null ? "-" : bgColor[0] + "_" + bgColor[1] + "_" + bgColor[2])
                + (required ? "_req" : "");
        return cache.computeIfAbsent(key, k -> {
            CellStyle style = wb.createCellStyle();
            style.cloneStyleFrom(baseStyle);
            // The mark is a number format, so the cell keeps the plain column name as its value
            if (required) style.setDataFormat(wb.createDataFormat().getFormat(REQUIRED_MARK_FORMAT));
            if (bgColor != null) {
                style.setFillForegroundColor(new XSSFColor(new byte[]{
                        (byte) bgColor[0], (byte) bgColor[1], (byte) bgColor[2]}));
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }
            if (fontColor == null) {
                return style;
            }
            Font font = wb.createFont();
            Font baseFont = wb.getFontAt(baseStyle.getFontIndex());
            font.setBold(baseFont.getBold());
            font.setFontHeight(baseFont.getFontHeight());
            font.setFontName(baseFont.getFontName());
            ((XSSFFont) font).setColor(
                    new XSSFColor(new byte[]{(byte) fontColor[0], (byte) fontColor[1], (byte) fontColor[2]}));
            style.setFont(font);
            return style;
        });
    }

    private static <T> void writeSingleHeaderRow(SXSSFSheet sheet, Cursor cursor,
                                                  List<ExcelColumn<T>> columns, CellStyle headerStyle,
                                                  @Nullable SXSSFWorkbook wb,
                                                  @Nullable Map<String, CellStyle> headerStyleCache,
                                                  float headerRowHeight) {
        SXSSFRow headRow = sheet.createRow(cursor.getRowOfSheet());
        if (headerRowHeight > 0) headRow.setHeightInPoints(headerRowHeight);
        cursor.plusRow();
        for (int j = 0; j < columns.size(); j++) {
            ExcelColumn<T> col = columns.get(j);
            SXSSFCell cell = headRow.createCell(j);
            ExcelCellText.write(cell, col.getName());
            cell.setCellStyle(resolveHeaderStyle(col, headerStyle, wb, headerStyleCache));
            applyHeaderComment(cell, col, sheet.getWorkbook());
        }
    }

    private static <T> void applyHeaderComment(SXSSFCell cell, ExcelColumn<T> col, SXSSFWorkbook wb) {
        ExcelCellComment hc = col.getHeaderComment();
        if (hc == null) return;
        int w = hc.width() > 0 ? hc.width() : col.getCommentWidth();
        int h = hc.height() > 0 ? hc.height() : col.getCommentHeight();
        ExcelWriteSupport.addCellComment(cell, hc.text(), hc.author(), w, h, wb);
    }

    /**
     * Writes N group header rows + 1 column header row.
     * <p>
     * For each column, its group levels (outermost→innermost) are top-aligned from the
     * first group row; missing levels at the bottom become {@code null} and merge
     * vertically into the column header cell.
     * Horizontal merges join adjacent columns with equal non-null values on the same row.
     */
    private static <T> void writeGroupAndColumnHeaders(SXSSFSheet sheet, Cursor cursor,
                                                        List<ExcelColumn<T>> columns, CellStyle headerStyle,
                                                        @Nullable SXSSFWorkbook wb,
                                                        @Nullable Map<String, CellStyle> headerStyleCache,
                                                        int maxDepth,
                                                        @Nullable Map<List<String>, ExcelCellComment> groupComments,
                                                        float headerRowHeight, boolean tableHeader) {
        int numCols = columns.size();
        int startRow = cursor.getRowOfSheet();

        SXSSFRow[] rows = createHeaderRows(sheet, cursor, maxDepth, headerRowHeight);
        String[][] grid = buildGroupGrid(columns, maxDepth);
        populateHeaderCells(rows, grid, columns, headerStyle, wb, headerStyleCache, maxDepth);
        applyHorizontalMerges(sheet, rows, grid, startRow, maxDepth, numCols, groupComments);
        applyVerticalMerges(sheet, rows, grid, columns, startRow, maxDepth, numCols, tableHeader);
    }

    /** Creates (maxDepth + 1) rows and advances the cursor. */
    private static SXSSFRow[] createHeaderRows(SXSSFSheet sheet, Cursor cursor,
                                                int maxDepth, float headerRowHeight) {
        SXSSFRow[] rows = new SXSSFRow[maxDepth + 1];
        for (int r = 0; r <= maxDepth; r++) {
            rows[r] = sheet.createRow(cursor.getRowOfSheet());
            if (headerRowHeight > 0) rows[r].setHeightInPoints(headerRowHeight);
            cursor.plusRow();
        }
        return rows;
    }

    /** Builds grid[depth][col] with group levels top-aligned. Unfilled slots are null. */
    private static <T> String[][] buildGroupGrid(List<ExcelColumn<T>> columns, int maxDepth) {
        String[][] grid = new String[maxDepth][columns.size()];
        for (int c = 0; c < columns.size(); c++) {
            String[] levels = columns.get(c).getGroupNames();
            for (int l = 0; l < levels.length; l++) {
                grid[l][c] = levels[l];
            }
        }
        return grid;
    }

    /** Styles and populates all header cells (group rows + column header row). */
    private static <T> void populateHeaderCells(SXSSFRow[] rows, String[][] grid,
                                                 List<ExcelColumn<T>> columns, CellStyle headerStyle,
                                                 @Nullable SXSSFWorkbook wb,
                                                 @Nullable Map<String, CellStyle> headerStyleCache,
                                                 int maxDepth) {
        for (int c = 0; c < columns.size(); c++) {
            ExcelColumn<T> col = columns.get(c);
            CellStyle colHeaderStyle = resolveHeaderStyle(col, headerStyle, wb, headerStyleCache);
            for (int r = 0; r < maxDepth; r++) {
                SXSSFCell cell = rows[r].createCell(c);
                cell.setCellStyle(colHeaderStyle);
                if (grid[r][c] != null) {
                    ExcelCellText.write(cell, grid[r][c]);
                }
            }
            SXSSFCell colCell = rows[maxDepth].createCell(c);
            colCell.setCellStyle(colHeaderStyle);
            ExcelCellText.write(colCell, col.getName());
        }
    }

    /** Merges adjacent columns with equal values per group row and attaches group comments. */
    private static void applyHorizontalMerges(SXSSFSheet sheet, SXSSFRow[] rows, String[][] grid,
                                               int startRow, int maxDepth, int numCols,
                                               @Nullable Map<List<String>, ExcelCellComment> groupComments) {
        for (int r = 0; r < maxDepth; r++) {
            int c = 0;
            while (c < numCols) {
                String v = grid[r][c];
                if (v == null) { c++; continue; }
                int start = c;
                // Equal labels merge only under the same parent groups.
                while (c < numCols && Objects.equals(v, grid[r][c]) && sameParents(grid, r, start, c)) c++;
                if (c - start > 1) {
                    sheet.addMergedRegion(new CellRangeAddress(startRow + r, startRow + r, start, c - 1));
                }
                if (groupComments != null && !groupComments.isEmpty()) {
                    List<String> path = new ArrayList<>(r + 1);
                    boolean valid = true;
                    for (int k = 0; k <= r; k++) {
                        String s = grid[k][start];
                        if (s == null) { valid = false; break; }
                        path.add(s);
                    }
                    if (valid) {
                        ExcelCellComment ec = groupComments.get(path);
                        if (ec != null) {
                            ExcelWriteSupport.addCellComment(rows[r].getCell(start), ec.text(), ec.author(),
                                    ec.width(), ec.height(), sheet.getWorkbook());
                        }
                    }
                }
            }
        }
    }

    private static boolean sameParents(String[][] grid, int level, int first, int other) {
        for (int k = 0; k < level; k++) {
            if (!Objects.equals(grid[k][first], grid[k][other])) return false;
        }
        return true;
    }

    /**
     * Merges trailing-null columns vertically into the column header cell.
     * Moves the column name to the top-left cell and blanks others to avoid bottom alignment.
     * <p>
     * Under a table nothing is merged into the column names row and its cells are never blank: Excel refuses a table
     * whose header cell is empty or part of a merged range. The name stays in the column names row and the cells
     * above it are left empty.
     */
    private static <T> void applyVerticalMerges(SXSSFSheet sheet, SXSSFRow[] rows, String[][] grid,
                                                  List<ExcelColumn<T>> columns,
                                                  int startRow, int maxDepth, int numCols, boolean tableHeader) {
        int columnHeaderRowIdx = startRow + maxDepth;
        for (int c = 0; c < numCols; c++) {
            int firstNullRow = maxDepth;
            for (int r = maxDepth - 1; r >= 0; r--) {
                if (grid[r][c] == null) firstNullRow = r;
                else break;
            }
            SXSSFCell topCell;
            if (firstNullRow < maxDepth && tableHeader) {
                // the name is already in the column names row; the cells above it stay empty and unmerged
                for (int r = firstNullRow; r < maxDepth; r++) rows[r].getCell(c).setBlank();
                topCell = rows[maxDepth].getCell(c);
            } else if (firstNullRow < maxDepth) {
                topCell = rows[firstNullRow].getCell(c);
                ExcelCellText.write(topCell, columns.get(c).getName());
                for (int r = firstNullRow + 1; r <= maxDepth; r++) {
                    rows[r].getCell(c).setBlank();
                }
                sheet.addMergedRegion(
                        new CellRangeAddress(startRow + firstNullRow, columnHeaderRowIdx, c, c));
            } else {
                topCell = rows[maxDepth].getCell(c);
            }
            applyHeaderComment(topCell, columns.get(c), sheet.getWorkbook());
        }
    }
}
