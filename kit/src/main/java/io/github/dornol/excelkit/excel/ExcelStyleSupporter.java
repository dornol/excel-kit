package io.github.dornol.excelkit.excel;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFFont;

import org.jspecify.annotations.Nullable;

import java.util.Arrays;
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
 * @since 2025-07-19
 */
class ExcelStyleSupporter {
    // Private constructor to prevent instantiation
    private ExcelStyleSupporter() {
        /* empty */
    }

    /**
     * Creates a bold, centered header cell style with a specified background color.
     * Automatically sets the font color to white if the background is dark.
     *
     * @param wb          SXSSFWorkbook instance
     * @param headerColor Background color of the header (XSSFColor)
     * @return Configured CellStyle for headers
     */
    static CellStyle headerStyle(SXSSFWorkbook wb, XSSFColor headerColor) {
        CellStyle headerStyle = wb.createCellStyle();
        Font headerFont = wb.createFont();
        headerStyle.setAlignment(HorizontalAlignment.CENTER);
        headerStyle.setVerticalAlignment(VerticalAlignment.CENTER);
        headerStyle.setFillForegroundColor(headerColor);
        headerStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        headerStyle.setBorderTop(BorderStyle.THIN);
        headerStyle.setBorderBottom(BorderStyle.THIN);
        headerStyle.setBorderLeft(BorderStyle.THIN);
        headerStyle.setBorderRight(BorderStyle.THIN);
        headerFont.setBold(true);
        headerFont.setFontHeight((short) (11 * 20));

        if (isDarkColor(headerColor)) {
            headerFont.setColor(HSSFColor.HSSFColorPredefined.WHITE.getIndex());
        } else {
            headerFont.setColor(HSSFColor.HSSFColorPredefined.BLACK.getIndex());
        }

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
        return luminance < 128; // threshold: below 128 out of 0-255 is considered dark
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
        return cellStyle(wb, alignment, format, null, null, null, null, null, cache);
    }

    static CellStyle cellStyle(SXSSFWorkbook wb, HorizontalAlignment alignment, @Nullable String format,
                               int @Nullable [] backgroundColor, @Nullable Boolean bold, @Nullable Integer fontSize,
                               @Nullable ExcelBorderStyle borderStyle, @Nullable Boolean locked,
                               Map<String, CellStyle> cache) {
        CellStyleParams params = new CellStyleParams(alignment, format, backgroundColor, bold, fontSize,
                borderStyle, locked, null, null, null, null, null, null, null, null);
        return cellStyle(wb, params, cache);
    }

    static CellStyle cellStyle(SXSSFWorkbook wb, CellStyleParams params, Map<String, CellStyle> cache) {
        String key = buildCacheKey(params);
        return cache.computeIfAbsent(key, k -> createCellStyle(wb, params));
    }

    private static String buildCacheKey(CellStyleParams params) {
        return params.alignment().name() + "|" + params.format()
                + "|" + Arrays.toString(params.backgroundColor())
                + "|" + params.bold() + "|" + params.fontSize()
                + "|" + params.borderStyle() + "|" + params.locked()
                + "|" + params.rotation()
                + "|" + params.borderTop() + "|" + params.borderBottom()
                + "|" + params.borderLeft() + "|" + params.borderRight()
                + "|" + Arrays.toString(params.fontColor())
                + "|" + params.strikethrough() + "|" + params.underline();
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

        // Font: bold, fontSize, fontColor, strikethrough, underline
        boolean needsFont = params.bold() != null || params.fontSize() != null
                || params.fontColor() != null || params.strikethrough() != null || params.underline() != null;
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
            nowStyle.setFont(font);
        }

        // Borders: per-side > uniform borderStyle > default THIN
        BorderStyle defaultBorder = (params.borderStyle() != null)
                ? params.borderStyle().toPoiBorderStyle() : BorderStyle.THIN;
        nowStyle.setBorderTop(params.borderTop() != null ? params.borderTop().toPoiBorderStyle() : defaultBorder);
        nowStyle.setBorderBottom(params.borderBottom() != null ? params.borderBottom().toPoiBorderStyle() : defaultBorder);
        nowStyle.setBorderLeft(params.borderLeft() != null ? params.borderLeft().toPoiBorderStyle() : defaultBorder);
        nowStyle.setBorderRight(params.borderRight() != null ? params.borderRight().toPoiBorderStyle() : defaultBorder);

        if (params.locked() != null) {
            nowStyle.setLocked(params.locked());
        }
        if (params.rotation() != null) {
            nowStyle.setRotation(params.rotation());
        }
        nowStyle.setWrapText(true);
        return nowStyle;
    }


}
