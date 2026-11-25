package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;

import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import org.jspecify.annotations.Nullable;

import java.util.Map;

/**
 * ExcelStyleSupporter
 * <p>
 * A utility class that provides reusable cell styles for header and body cells
 * in SXSSFWorkbook-based Excel exports.
 * </p>
 * This class supports dynamic font color adjustment for dark headers and border styling.
 *
 * @author dhkim
 * @since 0.1.2
 */
class ExcelStyleSupporter {
    // Private constructor to prevent instantiation
    private ExcelStyleSupporter() {
        /* empty */
    }

    private static final int DEFAULT_FONT_SIZE = 11;
    private static final int FONT_HEIGHT_MULTIPLIER = 20;
    private static final double LUMINANCE_DARK_THRESHOLD = 128;

    /**
     * Creates a bold, centered header cell style with a specified background color.
     * Automatically sets the font color to white if the background is dark.
     *
     * @param wb          SXSSFWorkbook instance
     * @param headerColor Background color of the header (XSSFColor)
     * @return Configured CellStyle for headers
     */
    static CellStyle headerStyle(SXSSFWorkbook wb, XSSFColor headerColor) {
        return headerStyle(wb, headerColor, null, null);
    }

    static CellStyle headerStyle(SXSSFWorkbook wb, XSSFColor headerColor,
                                  @Nullable String fontName, @Nullable Integer fontSize) {
        return headerStyle(wb, headerColor, fontName, fontSize, null);
    }

    static CellStyle headerStyle(SXSSFWorkbook wb, XSSFColor headerColor,
                                  @Nullable String fontName, @Nullable Integer fontSize,
                                  @Nullable HeaderStyleConfig hdrCfg) {
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();

        HorizontalAlignment hAlign = (hdrCfg != null && hdrCfg.alignment != null) ? hdrCfg.alignment : HorizontalAlignment.CENTER;
        VerticalAlignment vAlign = (hdrCfg != null && hdrCfg.verticalAlignment != null) ? hdrCfg.verticalAlignment : VerticalAlignment.CENTER;
        boolean bold = (hdrCfg != null && hdrCfg.bold != null) ? hdrCfg.bold : true;
        BorderStyle border = (hdrCfg != null && hdrCfg.borderStyle != null) ? toPoiBorderStyle(hdrCfg.borderStyle) : BorderStyle.THIN;

        headerStyle.setAlignment(hAlign);
        headerStyle.setVerticalAlignment(vAlign);
        headerStyle.setFillForegroundColor(headerColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderTop(border);
        headerStyle.setBorderBottom(border);
        headerStyle.setBorderLeft(border);
        headerStyle.setBorderRight(border);
        if (hdrCfg != null && hdrCfg.wrapText != null) {
            headerStyle.setWrapText(hdrCfg.wrapText);
        }
        headerFont.setBold(bold);
        headerFont.setFontHeight((short) ((fontSize != null ? fontSize : DEFAULT_FONT_SIZE) * FONT_HEIGHT_MULTIPLIER));
        if (fontName != null) {
            headerFont.setFontName(fontName);
        }

        // An RGB colour, not the palette's black or white: under a table Excel takes a palette colour for "automatic" and
        // shows the table style's own header text colour (white) instead, which vanishes on the white fill of the header.
        byte level = isDarkColor(headerColor) ? (byte) 0xFF : 0;
        ((XSSFFont) headerFont).setColor(new XSSFColor(new byte[]{level, level, level}));

        headerStyle.setFont(headerFont);
        return headerStyle;
    }

    /**
     * Determines whether the given XSSFColor is visually dark using luminance.
     * Uses W3C's brightness formula: Y = 0.299R + 0.587G + 0.114B
     *
     * @param color XSSFColor to evaluate
     * @return true if color is considered dark, false otherwise
     */
    private static boolean isDarkColor(XSSFColor color) {
        if (color == null) {
            return false;
        }
        byte[] rgb = color.getRGB();
        if (rgb == null || rgb.length != 3) {
            return false;
        }

        int r = Byte.toUnsignedInt(rgb[0]);
        int g = Byte.toUnsignedInt(rgb[1]);
        int b = Byte.toUnsignedInt(rgb[2]);

        // W3C brightness formula (Y = 0.299R + 0.587G + 0.114B)
        double luminance = 0.299 * r + 0.587 * g + 0.114 * b;
        return luminance < LUMINANCE_DARK_THRESHOLD;
    }

    /**
     * Returns a cached cell style for the given alignment and data format.
     * If no cached style exists for the combination, a new one is created and cached.
     *
     * @param wb        SXSSFWorkbook instance
     * @param alignment Cell horizontal alignment (e.g., CENTER, LEFT)
     * @param format    Data format string (e.g., "yyyy-mm-dd", "#,##0")
     * @param cache     Style cache keyed by alignment+format combination
     * @return Configured CellStyle for body cells
     */
    static CellStyle cellStyle(SXSSFWorkbook wb, HorizontalAlignment alignment, @Nullable String format, Map<String, CellStyle> cache) {
        return cellStyle(wb, CellStyleParams.of(alignment, format), cache);
    }

    static CellStyle cellStyle(SXSSFWorkbook wb, HorizontalAlignment alignment, @Nullable String format,
                               int @Nullable [] backgroundColor, @Nullable Boolean bold, @Nullable Integer fontSize,
                               @Nullable ExcelBorderStyle borderStyle, @Nullable Boolean locked,
                               Map<String, CellStyle> cache) {
        return cellStyle(wb, CellStyleParams.of(alignment, format, backgroundColor, bold, fontSize, borderStyle, locked), cache);
    }

    static CellStyle cellStyle(SXSSFWorkbook wb, CellStyleParams params, Map<String, CellStyle> cache) {
        String key = params.cacheKey();
        return cache.computeIfAbsent(key, k -> createCellStyle(wb, params));
    }

    private static CellStyle createCellStyle(SXSSFWorkbook wb, CellStyleParams params) {
        CellStyle nowStyle = wb.createCellStyle();

        nowStyle.setAlignment(params.alignment());
        if (params.format() != null) {
            DataFormat dataFormat = wb.createDataFormat();
            nowStyle.setDataFormat(dataFormat.getFormat(params.format()));
        }
        int[] backgroundColor = params.backgroundColor();
        if (backgroundColor != null) {
            nowStyle.setFillForegroundColor(new XSSFColor(new byte[]{
                    (byte) backgroundColor[0], (byte) backgroundColor[1], (byte) backgroundColor[2]}));
            nowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }

        // Font: bold, fontSize, fontColor, strikethrough, underline, fontName
        boolean needsFont = params.bold() != null || params.fontSize() != null
                || params.fontColor() != null || params.strikethrough() != null
                || params.underline() != null || params.fontName() != null;
        if (needsFont) {
            Font font = wb.createFont();
            if (params.bold() != null) {
                font.setBold(params.bold());
            }
            if (params.fontSize() != null) {
                font.setFontHeightInPoints(params.fontSize().shortValue());
            }
            int[] fontColor = params.fontColor();
            if (fontColor != null) {
                ((XSSFFont) font).setColor(new XSSFColor(new byte[]{
                        (byte) fontColor[0], (byte) fontColor[1], (byte) fontColor[2]}));
            }
            if (params.strikethrough() != null) {
                font.setStrikeout(params.strikethrough());
            }
            if (params.underline() != null && params.underline()) {
                font.setUnderline(Font.U_SINGLE);
            }
            if (params.fontName() != null) {
                font.setFontName(params.fontName());
            }
            nowStyle.setFont(font);
        }

        // Borders: per-side > uniform borderStyle > default THIN
        BorderStyle defaultBorder = (params.borderStyle() != null)
                ? toPoiBorderStyle(params.borderStyle()) : BorderStyle.THIN;
        nowStyle.setBorderTop(params.borderTop() != null ? toPoiBorderStyle(params.borderTop()) : defaultBorder);
        nowStyle.setBorderBottom(params.borderBottom() != null ? toPoiBorderStyle(params.borderBottom()) : defaultBorder);
        nowStyle.setBorderLeft(params.borderLeft() != null ? toPoiBorderStyle(params.borderLeft()) : defaultBorder);
        nowStyle.setBorderRight(params.borderRight() != null ? toPoiBorderStyle(params.borderRight()) : defaultBorder);

        if (params.locked() != null) {
            nowStyle.setLocked(params.locked());
        }
        if (params.rotation() != null) {
            nowStyle.setRotation(params.rotation());
        }
        nowStyle.setVerticalAlignment(params.verticalAlignment() != null
                ? params.verticalAlignment() : VerticalAlignment.CENTER);
        nowStyle.setWrapText(params.wrapText() != null ? params.wrapText() : true);
        if (params.indentation() != null) {
            nowStyle.setIndention(params.indentation());
        }
        return nowStyle;
    }



    /**
     * Maps an {@link ExcelBorderStyle} to the Apache POI border style of the same name.
     */
    static BorderStyle toPoiBorderStyle(ExcelBorderStyle style) {
        return BorderStyle.valueOf(style.name());
    }

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
        return resolveRowStyle(baseStyle, bgColor, rowStyle, cache, wb, maxStyles, null);
    }

    /** @param dataFormat the number format to give the style instead of the base style's, or {@code null} to keep it */
    static CellStyle resolveRowStyle(CellStyle baseStyle, @Nullable ExcelColor bgColor,
                                      RowStyleConfig rowStyle, Map<DerivedStyleKey, CellStyle> cache,
                                      SXSSFWorkbook wb, int maxStyles, @Nullable String dataFormat) {
        DerivedStyleKey key = new DerivedStyleKey(baseStyle.getIndex(), DerivedStyleKey.rgb(bgColor),
                dataFormat == null ? rowStyle.cacheKey() : rowStyle.cacheKey() + "|fmt=" + dataFormat);
        return cache.computeIfAbsent(key, k -> {
            CellStyle style = createDerivedStyle(wb, maxStyles);
            style.cloneStyleFrom(baseStyle);
            if (dataFormat != null) style.setDataFormat(wb.createDataFormat().getFormat(dataFormat));

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

    static <T> CellStyle emptyMessageStyle(SXSSFWorkbook wb, SheetConfig<T> cfg) {
        CellStyle created = createDerivedStyle(wb, cfg.maxCellStyles);
        XSSFFont font = (XSSFFont) wb.createFont();
        font.setItalic(true);
        font.setColor(new XSSFColor(new byte[]{(byte) 128, (byte) 128, (byte) 128}));
        created.setFont(font);
        created.setAlignment(HorizontalAlignment.CENTER);
        created.setVerticalAlignment(VerticalAlignment.CENTER);
        return created;
    }
}
