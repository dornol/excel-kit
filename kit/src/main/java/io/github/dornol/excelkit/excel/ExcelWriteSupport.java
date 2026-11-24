package io.github.dornol.excelkit.excel;

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
        return writeAfterDataAndSummary(sheet, wb, startRow, columns, headerRowIndex, cfg, null);
    }

    /** @param values what the summary formulas calculate for the sheet, or {@code null} when they are not tracked */
    static <T> int writeAfterDataAndSummary(SXSSFSheet sheet, SXSSFWorkbook wb, int startRow,
                                             List<ExcelColumn<T>> columns, int headerRowIndex,
                                             SheetConfig<T> cfg, @Nullable FormulaValues values) {
        return writeAfterDataAndSummary(sheet, wb, startRow, columns, headerRowIndex, cfg, values, startRow);
    }

    /**
     * @param dataEndRow the 1-based last data row, which is where the summary ranges end: {@code startRow} when the
     *                   rows below the data are the next ones, less when something that is not data (the empty
     *                   message) sits between
     */
    static <T> int writeAfterDataAndSummary(SXSSFSheet sheet, SXSSFWorkbook wb, int startRow,
                                             List<ExcelColumn<T>> columns, int headerRowIndex,
                                             SheetConfig<T> cfg, @Nullable FormulaValues values, int dataEndRow) {
        int row = startRow;
        if (cfg.afterDataWriter != null) {
            row = cfg.afterDataWriter.write(new SheetContext(sheet, wb, row, columns, headerRowIndex));
        }
        if (cfg.summaryConfig != null) {
            // The summary covers the data rows only, not whatever afterData wrote below them.
            row = cfg.summaryConfig.toAfterDataWriter(cfg.subtotalKey != null, values).write(
                    new SheetContext(sheet, wb, row, columns, headerRowIndex, dataEndRow));
        }
        return row;
    }

    /** Rejects what the empty message would corrupt: a merged line has no place in a table. */
    static <T> void validateEmptyMessage(SheetConfig<T> cfg) {
        if (cfg.emptyMessage != null && cfg.tableOptions != null) {
            throw new ExcelWriteException("emptyMessage(...) cannot be combined with table(...): the message is a merged "
                    + "cell, and Excel does not allow merged cells in a table");
        }
    }

    private static final String EMPTY_MESSAGE_STYLE = "excel-kit:empty-message";

    private static <T> CellStyle emptyMessageStyle(SXSSFWorkbook wb, SheetConfig<T> cfg) {
        CellStyle created = createDerivedStyle(wb, cfg.maxCellStyles);
        XSSFFont font = (XSSFFont) wb.createFont();
        font.setItalic(true);
        font.setColor(new XSSFColor(new byte[]{(byte) 128, (byte) 128, (byte) 128}));
        created.setFont(font);
        created.setAlignment(HorizontalAlignment.CENTER);
        created.setVerticalAlignment(VerticalAlignment.CENTER);
        return created;
    }

    /**
     * Writes the line of an empty sheet below the header: the text centered over all the columns, in gray italics.
     * It takes a row of the sheet but is no data row.
     */
    static <T> void writeEmptyMessage(SXSSFSheet sheet, SXSSFWorkbook wb, Cursor cursor, List<ExcelColumn<T>> columns,
                                      SheetConfig<T> cfg, Map<DerivedStyleKey, CellStyle> styleCache,
                                      @Nullable Map<String, CellStyle> workbookStyles) {
        String text = cfg.emptyMessage;
        if (text == null) return;
        int rowIndex = cursor.getRowOfSheet();
        if (rowIndex >= ExcelLimits.SHEET_ROW_CAPACITY) {
            throw new ExcelWriteException("Sheet '" + sheet.getSheetName() + "' is full: no room for the empty message");
        }
        // The style does not depend on the columns, so the sheets of a workbook share one.
        CellStyle style = workbookStyles != null
                ? workbookStyles.computeIfAbsent(EMPTY_MESSAGE_STYLE, k -> emptyMessageStyle(wb, cfg))
                : styleCache.computeIfAbsent(new DerivedStyleKey(-1, -1, "empty-message"), k -> emptyMessageStyle(wb, cfg));
        SXSSFRow row = sheet.createRow(rowIndex);
        row.setHeightInPoints(cfg.rowHeightInPoints);
        cursor.plusRow();
        for (int i = 0; i < columns.size(); i++) {
            SXSSFCell cell = row.createCell(i);
            if (i == 0) ExcelCellText.write(cell, text);
            cell.setCellStyle(style);
        }
        if (columns.size() > 1) sheet.addMergedRegion(new CellRangeAddress(rowIndex, rowIndex, 0, columns.size() - 1));
    }

    /** What a sheet needs as soon as its header is written. The filter waits for the data: see {@code ExcelSheetPostProcessor}. */
    static void applySheetOptions(SXSSFSheet sheet, int headerRowIdx, int freezePaneCols, int freezePaneRows) {
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

    /**
     * Identifies a style derived from a column's base style: the base style, the background color
     * ({@code -1} for none) and, for a row style, its {@link RowStyleConfig#cacheKey()}. A record keeps the
     * per-cell lookup free of string building.
     */
    record DerivedStyleKey(int baseIndex, int rgb, @Nullable String rule) {
        static int rgb(@Nullable ExcelColor color) {
            return color == null ? -1 : (color.getR() << 16) | (color.getG() << 8) | color.getB();
        }
    }

    static CellStyle resolveColorStyle(CellStyle baseStyle, ExcelColor color,
                                        Map<DerivedStyleKey, CellStyle> cache, SXSSFWorkbook wb, int maxStyles) {
        DerivedStyleKey key = new DerivedStyleKey(baseStyle.getIndex(), DerivedStyleKey.rgb(color), null);
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
                                      RowStyleConfig rowStyle, Map<DerivedStyleKey, CellStyle> cache,
                                      SXSSFWorkbook wb, int maxStyles) {
        DerivedStyleKey key = new DerivedStyleKey(baseStyle.getIndex(), DerivedStyleKey.rgb(bgColor),
                rowStyle.cacheKey());
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
            if (rowStyle.locked != null) style.setLocked(rowStyle.locked);
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

    /**
     * Rejects a column that merges equal values where the sheet cannot hold merged cells: in a table (Excel repairs
     * the file) or in a list written into a template.
     */
    static <T> void validateMergeColumns(List<ExcelColumn<T>> columns, @Nullable TableOptions tableOptions, boolean supported) {
        for (ExcelColumn<T> col : columns) {
            if (!col.isMergeSameValues()) continue;
            if (!supported) {
                throw new ExcelWriteException("Column '" + col.getName() + "' asks for mergeSameValues(), "
                        + "which a template list writer does not apply");
            }
            if (tableOptions != null) {
                throw new ExcelWriteException("Column '" + col.getName() + "' asks for mergeSameValues(), but "
                        + "the sheet is a table(...): Excel does not allow merged cells in a table");
            }
        }
    }

    static <T> void validateUniqueColumnNames(List<ExcelColumn<T>> columns) {
        Set<String> seen = new HashSet<>();
        for (ExcelColumn<T> col : columns) {
            if (!seen.add(col.getName())) {
                throw new ExcelWriteException("Duplicate column name: '" + col.getName() + "'");
            }
        }
    }

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
        CellStyle style = config != null && config.customStyle != null
                ? config.customStyle
                : ExcelStyleSupporter.cellStyle(wb, c.styleParams(), cellStyleCache);

        return new ExcelColumn<>(name, function, style,
                dataType != null ? dataType.getSetter() : ExcelDataType.STRING.getSetter(), c)
                .sharingStylesOf(cellStyleCache);
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

    private static final int MAX_SHEET_NAME_LENGTH = 31;

    /**
     * Name of the {@code index}-th sheet (0-based) of a rolled-over write: the first keeps the base
     * name, later ones get a " (n)" suffix. Excel limits sheet names to 31 characters, so the base is
     * shortened to keep the suffix and the names unique.
     */
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

    /**
     * Hides the sheet. A hidden sheet is not the one the workbook opens on: it is also deselected, since a selected
     * hidden sheet would make Excel type into it together with the visible one.
     */
    static void applyVisibility(SXSSFSheet sheet, org.apache.poi.ss.usermodel.@Nullable SheetVisibility visibility) {
        if (visibility == null || visibility == org.apache.poi.ss.usermodel.SheetVisibility.VISIBLE) return;
        var wb = sheet.getWorkbook();
        wb.setSheetVisibility(wb.getSheetIndex(sheet), visibility);
        SXSSFSheetHelper.getXSSFSheet(sheet).setSelected(false);
    }

    /**
     * Checks that a workbook keeps a visible sheet (Excel cannot open one with none) and that it opens on one.
     *
     * @throws ExcelWriteException if every sheet is hidden
     */
    static void requireVisibleSheet(org.apache.poi.xssf.streaming.SXSSFWorkbook wb) {
        int first = -1;
        for (int i = 0; i < wb.getNumberOfSheets(); i++) {
            if (wb.getSheetVisibility(i) == org.apache.poi.ss.usermodel.SheetVisibility.VISIBLE) {
                first = i;
                break;
            }
        }
        if (first < 0) {
            if (wb.getNumberOfSheets() == 0) return;
            throw new ExcelWriteException("Every sheet of the workbook is hidden; Excel cannot open a workbook "
                    + "without a visible sheet, so leave at least one sheet visible");
        }
        if (wb.getSheetVisibility(wb.getActiveSheetIndex()) != org.apache.poi.ss.usermodel.SheetVisibility.VISIBLE) {
            wb.setActiveSheet(first);
            SXSSFSheetHelper.getXSSFSheet(wb.getSheetAt(first)).setSelected(true);
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
                                  int headerRowIndex, int dataEndRow, int sheetIndex) {
        if (namedRanges == null || namedRanges.isEmpty()) return;
        int lastRow = sheet.getLastRowNum();
        if (dataEndRow <= headerRowIndex) return; // no data rows
        String sheetName = sheet.getSheetName();
        var wb = sheet.getWorkbook();
        for (var entry : namedRanges.entrySet()) {
            String col = SheetContext.columnLetter(entry.getValue());
            int dataStart = headerRowIndex + 2; // 1-based, after header
            // Concatenation, not String.format: %d writes the digits of the default locale (Arabic-Indic, ...)
            String ref = "'" + sheetName.replace("'", "''") + "'!$" + col + "$" + dataStart + ":$" + col + "$" + (lastRow + 1);
            var name = wb.createName();
            name.setNameName(sheetIndex == 0 ? entry.getKey() : entry.getKey() + "_" + (sheetIndex + 1));
            name.setRefersToFormula(ref);
        }
    }

}
