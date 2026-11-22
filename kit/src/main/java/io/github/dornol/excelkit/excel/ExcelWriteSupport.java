package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelCellComment;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import io.github.dornol.excelkit.core.Cursor;
import io.github.dornol.excelkit.core.ProgressCallback;
import io.github.dornol.excelkit.core.read.RowFunction;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFFont;

import org.jspecify.annotations.Nullable;

import java.util.ArrayList;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.Set;

/**
 * Package-private utility methods shared by {@link ExcelWriter} and {@link ExcelSheetWriter}
 * to eliminate duplicate write logic.
 *
 * @author dhkim
 */
class ExcelWriteSupport {

    /**
     * Creates a style for a color or row style, or says why the workbook cannot take another one:
     * {@code maxStyles} is either the .xlsx limit or a lower one the caller asked for.
     */
    static CellStyle createDerivedStyle(SXSSFWorkbook wb, int maxStyles) {
        if (wb.getNumCellStyles() >= maxStyles) {
            String limit = maxStyles >= ExcelLimits.MAX_CELL_STYLES
                    ? "an .xlsx workbook holds at most " + ExcelLimits.MAX_CELL_STYLES
                    : "maxCellStyles(" + maxStyles + ") allows at most " + maxStyles;
            throw new ExcelWriteException("Too many distinct cell styles: " + limit
                    + ". Each distinct color returned by cellColor, rowColor or rowStyle needs its own style; "
                    + "return fewer distinct colors, for example by rounding values into a small palette.");
        }
        return wb.createCellStyle();
    }


    private ExcelWriteSupport() {
    }

    /**
     * Invokes the afterData and summary callbacks on the given sheet, returning the next
     * available row index. Used at both rollover points and post-data finalization.
     */
    static <T> int writeAfterDataAndSummary(SXSSFSheet sheet, SXSSFWorkbook wb, int startRow,
                                             List<ExcelColumn<T>> columns, int headerRowIndex,
                                             SheetConfig<T> cfg) {
        int row = startRow;
        if (cfg.afterDataWriter != null) {
            row = cfg.afterDataWriter.write(new SheetContext(sheet, wb, row, columns, headerRowIndex));
        }
        if (cfg.summaryConfig != null) {
            // The summary covers the data rows only, not whatever afterData wrote below them.
            row = cfg.summaryConfig.toAfterDataWriter().write(
                    new SheetContext(sheet, wb, row, columns, headerRowIndex, startRow));
        }
        return row;
    }

    /**
     * Writes column headers, with 0..N optional group header rows if any column has groups.
     */
    static <T> void writeColumnHeaders(SXSSFSheet sheet, Cursor cursor,
                                        List<ExcelColumn<T>> columns, CellStyle headerStyle) {
        writeColumnHeaders(sheet, cursor, columns, headerStyle, null, null, null, 0f);
    }

    static <T> void writeColumnHeaders(SXSSFSheet sheet, Cursor cursor,
                                        List<ExcelColumn<T>> columns, CellStyle headerStyle,
                                        @Nullable SXSSFWorkbook wb, @Nullable Map<String, CellStyle> headerStyleCache) {
        writeColumnHeaders(sheet, cursor, columns, headerStyle, wb, headerStyleCache, null, 0f);
    }

    static <T> void writeColumnHeaders(SXSSFSheet sheet, Cursor cursor,
                                        List<ExcelColumn<T>> columns, CellStyle headerStyle,
                                        @Nullable SXSSFWorkbook wb, @Nullable Map<String, CellStyle> headerStyleCache,
                                        @Nullable Map<List<String>, ExcelCellComment> groupComments,
                                        float headerRowHeight) {
        int maxDepth = 0;
        for (ExcelColumn<T> c : columns) {
            int d = c.getGroupNames().length;
            if (d > maxDepth) maxDepth = d;
        }
        if (maxDepth == 0) {
            writeSingleHeaderRow(sheet, cursor, columns, headerStyle, wb, headerStyleCache, headerRowHeight);
        } else {
            writeGroupAndColumnHeaders(sheet, cursor, columns, headerStyle, wb, headerStyleCache,
                    maxDepth, groupComments, headerRowHeight);
        }
    }

    private static <T> CellStyle resolveHeaderStyle(ExcelColumn<T> col, CellStyle baseStyle,
                                                     @Nullable SXSSFWorkbook wb,
                                                     @Nullable Map<String, CellStyle> cache) {
        int[] fontColor = col.getHeaderFontColor();
        int[] bgColor = col.getHeaderBackgroundColor();
        if ((fontColor == null && bgColor == null) || wb == null || cache == null) {
            return baseStyle;
        }
        String key = "hdr_" + baseStyle.getIndex()
                + "_f" + (fontColor == null ? "-" : fontColor[0] + "_" + fontColor[1] + "_" + fontColor[2])
                + "_b" + (bgColor == null ? "-" : bgColor[0] + "_" + bgColor[1] + "_" + bgColor[2]);
        return cache.computeIfAbsent(key, k -> {
            CellStyle style = wb.createCellStyle();
            style.cloneStyleFrom(baseStyle);
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
        addCellComment(cell, hc.text(), hc.author(), w, h, wb);
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
                                                        float headerRowHeight) {
        int numCols = columns.size();
        int startRow = cursor.getRowOfSheet();

        SXSSFRow[] rows = createHeaderRows(sheet, cursor, maxDepth, headerRowHeight);
        String[][] grid = buildGroupGrid(columns, maxDepth);
        populateHeaderCells(rows, grid, columns, headerStyle, wb, headerStyleCache, maxDepth);
        applyHorizontalMerges(sheet, rows, grid, startRow, maxDepth, numCols, groupComments);
        applyVerticalMerges(sheet, rows, grid, columns, startRow, maxDepth, numCols);
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
                            addCellComment(rows[r].getCell(start), ec.text(), ec.author(),
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
     */
    private static <T> void applyVerticalMerges(SXSSFSheet sheet, SXSSFRow[] rows, String[][] grid,
                                                  List<ExcelColumn<T>> columns,
                                                  int startRow, int maxDepth, int numCols) {
        int columnHeaderRowIdx = startRow + maxDepth;
        for (int c = 0; c < numCols; c++) {
            int firstNullRow = maxDepth;
            for (int r = maxDepth - 1; r >= 0; r--) {
                if (grid[r][c] == null) firstNullRow = r;
                else break;
            }
            SXSSFCell topCell;
            if (firstNullRow < maxDepth) {
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

    static void applySheetOptions(SXSSFSheet sheet, int headerRowIdx,
                                   boolean autoFilter, int freezePaneCols, int freezePaneRows, int columnCount) {
        if (autoFilter) {
            sheet.setAutoFilter(new CellRangeAddress(headerRowIdx, headerRowIdx, 0, columnCount - 1));
        }
        if (freezePaneCols > 0 || freezePaneRows > 0) {
            sheet.createFreezePane(freezePaneCols, headerRowIdx + freezePaneRows);
        }
    }

    static void addCellComment(SXSSFCell cell, String text, SXSSFWorkbook wb) {
        addCellComment(cell, text, null, 0, 0, wb);
    }

    static void addCellComment(SXSSFCell cell, String text, @Nullable String author,
                                int width, int height, SXSSFWorkbook wb) {
        Drawing<?> drawing = cell.getSheet().createDrawingPatriarch();
        CreationHelper factory = wb.getCreationHelper();
        ClientAnchor anchor = factory.createClientAnchor();
        int w = width > 0 ? width : 2;
        int h = height > 0 ? height : 3;
        anchor.setCol1(cell.getColumnIndex());
        anchor.setCol2(cell.getColumnIndex() + w);
        anchor.setRow1(cell.getRowIndex());
        anchor.setRow2(cell.getRowIndex() + h);
        Comment comment = drawing.createCellComment(anchor);
        comment.setString(factory.createRichTextString(text));
        if (author != null) {
            comment.setAuthor(author);
        }
        cell.setCellComment(comment);
    }

    static void applyChart(SXSSFSheet sheet, @Nullable ExcelChartConfig chartConfig,
                            int headerRow, int dataEndRow, int nextFreeRow, int columnCount) {
        if (chartConfig != null) {
            chartConfig.apply(sheet, headerRow, dataEndRow, nextFreeRow, columnCount);
        }
    }

    static CellStyle resolveColorStyle(CellStyle baseStyle, ExcelColor color,
                                        Map<String, CellStyle> cache, SXSSFWorkbook wb, int maxStyles) {
        String key = baseStyle.getIndex() + "_" + color.getR() + "_" + color.getG() + "_" + color.getB();
        return cache.computeIfAbsent(key, k -> {
            CellStyle style = createDerivedStyle(wb, maxStyles);
            style.cloneStyleFrom(baseStyle);
            style.setFillForegroundColor(new XSSFColor(new byte[]{
                    (byte) color.getR(), (byte) color.getG(), (byte) color.getB()}));
            style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            return style;
        });
    }

    /**
     * A new font with every attribute of {@code base}, so that a row style changing one of them
     * (bold, size, ...) does not drop the others, such as an RGB font color or an underline.
     */
    private static Font copyFont(Font base, SXSSFWorkbook wb) {
        Font copy = wb.createFont();
        copy.setFontName(base.getFontName());
        copy.setFontHeight(base.getFontHeight());
        copy.setBold(base.getBold());
        copy.setItalic(base.getItalic());
        copy.setStrikeout(base.getStrikeout());
        copy.setUnderline(base.getUnderline());
        copy.setTypeOffset(base.getTypeOffset());
        copy.setCharSet(base.getCharSet());
        copy.setColor(base.getColor());
        if (base instanceof XSSFFont from && copy instanceof XSSFFont to) {
            if (from.getXSSFColor() != null) to.setColor(from.getXSSFColor());
            to.setFamily(from.getFamily());
        }
        return copy;
    }

    static CellStyle resolveRowStyle(CellStyle baseStyle, @Nullable ExcelColor bgColor,
                                      RowStyleConfig rowStyle, Map<String, CellStyle> cache,
                                      SXSSFWorkbook wb, int maxStyles) {
        String bgKey = bgColor != null
                ? bgColor.getR() + "_" + bgColor.getG() + "_" + bgColor.getB()
                : "none";
        String key = baseStyle.getIndex() + "_" + rowStyle.cacheKey() + "_bg" + bgKey;
        return cache.computeIfAbsent(key, k -> {
            CellStyle style = createDerivedStyle(wb, maxStyles);
            style.cloneStyleFrom(baseStyle);

            // Background color
            if (bgColor != null) {
                style.setFillForegroundColor(new XSSFColor(new byte[]{
                        (byte) bgColor.getR(), (byte) bgColor.getG(), (byte) bgColor.getB()}));
                style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
            }

            // Font modifications
            if (rowStyle.bold != null || rowStyle.fontSize != null || rowStyle.fontColor != null
                    || rowStyle.italic != null || rowStyle.strikethrough != null) {
                Font newFont = copyFont(wb.getFontAt(baseStyle.getFontIndex()), wb);

                if (rowStyle.bold != null) newFont.setBold(rowStyle.bold);
                if (rowStyle.fontSize != null) newFont.setFontHeightInPoints(rowStyle.fontSize.shortValue());
                if (rowStyle.italic != null) newFont.setItalic(rowStyle.italic);
                if (rowStyle.strikethrough != null) newFont.setStrikeout(rowStyle.strikethrough);
                if (rowStyle.fontColor != null) {
                    newFont.setColor(XSSFFont.DEFAULT_FONT_COLOR);
                    if (newFont instanceof XSSFFont xf) {
                        xf.setColor(new XSSFColor(new byte[]{
                                (byte) rowStyle.fontColor.getR(),
                                (byte) rowStyle.fontColor.getG(),
                                (byte) rowStyle.fontColor.getB()}));
                    }
                }
                style.setFont(newFont);
            }
            return style;
        });
    }

    static <T> int initSheetPreamble(SXSSFSheet sheet, SXSSFWorkbook wb,
                                      List<ExcelColumn<T>> columns,
                                      @Nullable BeforeHeaderWriter writer) {
        int currentRow = 0;
        if (writer != null) {
            currentRow = writer.write(new SheetContext(sheet, wb, currentRow, columns));
        }
        return currentRow;
    }

    static <T> void validateUniqueColumnNames(List<ExcelColumn<T>> columns) {
        Set<String> seen = new HashSet<>();
        for (ExcelColumn<T> col : columns) {
            if (!seen.add(col.getName())) {
                throw new ExcelWriteException("Duplicate column name: '" + col.getName() + "'");
            }
        }
    }

    private static final int MAX_SHEET_NAME_LENGTH = 31;

    /**
     * Name of the {@code index}-th sheet (0-based) of a rolled-over write: the first keeps the base
     * name, later ones get a " (n)" suffix. Excel limits sheet names to 31 characters, so the base is
     * shortened to keep the suffix and the names unique.
     */
    /**
     * Builds a column from a function and an optional configuration: applies the sheet's default style, resolves
     * the data format and the cell style (shared through {@code cellStyleCache}) and picks the setter of the data
     * type, which is plain text when none was given.
     */
    static <T> ExcelColumn<T> buildColumn(SXSSFWorkbook wb, Map<String, CellStyle> cellStyleCache,
                                          ColumnStyleConfig.@Nullable DefaultStyleConfig<T> defaultStyle, String name,
                                          RowFunction<T, @Nullable Object> function, @Nullable ColumnConfig<T> config) {
        Objects.requireNonNull(function, "function cannot be null");
        ColumnStyleConfig<T, ?> c = config != null ? config : new ColumnConfig<>();
        if (defaultStyle != null) {
            c.applyDefaults(defaultStyle);
        }
        ExcelDataType dataType = c.dataType;
        String dataFormat = c.dataFormat != null ? c.dataFormat
                : (dataType != null ? dataType.getDefaultFormat() : null);

        CellStyleParams params = new CellStyleParams(c.alignment, dataFormat,
                c.backgroundColor, c.bold, c.fontSize, c.borderStyle, c.locked,
                c.rotation, c.borderTop, c.borderBottom, c.borderLeft, c.borderRight,
                c.fontColor, c.strikethrough, c.underline,
                c.verticalAlignment, c.wrapText, c.fontName, c.indentation);
        CellStyle style = ExcelStyleSupporter.cellStyle(wb, params, cellStyleCache);

        return new ExcelColumn<>(name, function, style,
                dataType != null ? dataType.getSetter() : ExcelDataType.STRING.getSetter(), c);
    }

    /**
     * Rejects a sheet name Excel cannot hold, so a mistake surfaces where the name is given instead of
     * when the first sheet is created. The length is not checked here: a long base name is shortened for
     * the sheets that roll over.
     */
    static String requireValidSheetName(String name) {
        if (name.isEmpty()) {
            throw new IllegalArgumentException("Sheet name must not be empty");
        }
        for (int i = 0; i < name.length(); i++) {
            if ("\\/?*[]:".indexOf(name.charAt(i)) >= 0) {
                throw new IllegalArgumentException("Invalid char (" + name.charAt(i) + ") found at index (" + i
                        + ") in sheet name '" + name + "'");
            }
        }
        if (name.charAt(0) == '\'' || name.charAt(name.length() - 1) == '\'') {
            throw new IllegalArgumentException("Invalid sheet name '" + name
                    + "'. Sheet names must not begin or end with (').");
        }
        return name;
    }

    static String rolloverSheetName(String baseName, int index) {
        if (index == 0) return baseName;
        String suffix = " (" + (index + 1) + ")";
        int baseLength = Math.min(baseName.length(), Math.max(0, MAX_SHEET_NAME_LENGTH - suffix.length()));
        if (baseLength > 0 && baseLength < baseName.length()
                && Character.isHighSurrogate(baseName.charAt(baseLength - 1))) {
            baseLength--; // do not split a surrogate pair
        }
        return baseName.substring(0, baseLength) + suffix;
    }

    static void checkProgress(Cursor cursor, int interval, @Nullable ProgressCallback callback) {
        if (callback != null && interval > 0 && cursor.getCurrentTotal() % interval == 0) {
            callback.onProgress(cursor.getCurrentTotal(), cursor);
        }
    }

    static void applyTabColor(SXSSFSheet sheet, int @Nullable [] tabColor) {
        if (tabColor == null) return;
        XSSFSheet xssfSheet = SXSSFSheetHelper.getXSSFSheet(sheet);
        xssfSheet.setTabColor(new XSSFColor(new byte[]{
                (byte) tabColor[0], (byte) tabColor[1], (byte) tabColor[2]}));
    }

    /**
     * Defines the named ranges of one sheet. A workbook cannot hold two names alike, so the sheets of a
     * rollover get {@code name_2}, {@code name_3}, ... (the first keeps {@code name}), as tables do.
     */
    static void applyNamedRanges(SXSSFSheet sheet, @Nullable Map<String, Integer> namedRanges,
                                  int headerRowIndex, int sheetIndex) {
        if (namedRanges == null || namedRanges.isEmpty()) return;
        int lastRow = sheet.getLastRowNum();
        if (lastRow <= headerRowIndex) return; // no data rows
        String sheetName = sheet.getSheetName();
        var wb = sheet.getWorkbook();
        for (var entry : namedRanges.entrySet()) {
            String col = SheetContext.columnLetter(entry.getValue());
            int dataStart = headerRowIndex + 2; // 1-based, after header
            String ref = "'%s'!$%s$%d:$%s$%d".formatted(sheetName.replace("'", "''"), col, dataStart, col, lastRow + 1);
            var name = wb.createName();
            name.setNameName(sheetIndex == 0 ? entry.getKey() : entry.getKey() + "_" + (sheetIndex + 1));
            name.setRefersToFormula(ref);
        }
    }

}
