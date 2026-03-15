package io.github.dornol.excelkit.excel;

import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFColor;

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

        // ✅ W3C 표준 밝기 공식 (Y = 0.299R + 0.587G + 0.114B)
        double luminance = 0.299 * r + 0.587 * g + 0.114 * b;
        return luminance < 128; // 밝기 기준: 0~255 중 128 미만은 어둡다고 판단
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
    static CellStyle cellStyle(SXSSFWorkbook wb, HorizontalAlignment alignment, String format, Map<String, CellStyle> cache) {
        return cellStyle(wb, alignment, format, null, null, null, null, null, cache);
    }

    static CellStyle cellStyle(SXSSFWorkbook wb, HorizontalAlignment alignment, String format,
                               int[] backgroundColor, Boolean bold, Integer fontSize,
                               ExcelBorderStyle borderStyle, Boolean locked,
                               Map<String, CellStyle> cache) {
        StringBuilder keyBuilder = new StringBuilder();
        keyBuilder.append(alignment.name()).append("|").append(format);
        if (backgroundColor != null) {
            keyBuilder.append("|bg=").append(backgroundColor[0]).append(",").append(backgroundColor[1]).append(",").append(backgroundColor[2]);
        }
        if (bold != null) {
            keyBuilder.append("|bold=").append(bold);
        }
        if (fontSize != null) {
            keyBuilder.append("|fs=").append(fontSize);
        }
        if (borderStyle != null) {
            keyBuilder.append("|border=").append(borderStyle.name());
        }
        if (locked != null) {
            keyBuilder.append("|locked=").append(locked);
        }
        String key = keyBuilder.toString();
        return cache.computeIfAbsent(key, k -> createCellStyle(wb, alignment, format, backgroundColor, bold, fontSize, borderStyle, locked));
    }

    private static CellStyle createCellStyle(SXSSFWorkbook wb, HorizontalAlignment alignment, String format,
                                             int[] backgroundColor, Boolean bold, Integer fontSize,
                                             ExcelBorderStyle borderStyle, Boolean locked) {
        CellStyle nowStyle = wb.createCellStyle();

        nowStyle.setAlignment(alignment);
        if (format != null) {
            DataFormat dataFormat = wb.createDataFormat();
            nowStyle.setDataFormat(dataFormat.getFormat(format));
        }
        if (backgroundColor != null) {
            nowStyle.setFillForegroundColor(new XSSFColor(new byte[]{
                    (byte) backgroundColor[0], (byte) backgroundColor[1], (byte) backgroundColor[2]}));
            nowStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
        }
        if (bold != null || fontSize != null) {
            Font font = wb.createFont();
            if (bold != null) {
                font.setBold(bold);
            }
            if (fontSize != null) {
                font.setFontHeightInPoints(fontSize.shortValue());
            }
            nowStyle.setFont(font);
        }
        BorderStyle border = (borderStyle != null) ? borderStyle.toPoiBorderStyle() : BorderStyle.THIN;
        nowStyle.setBorderTop(border);
        nowStyle.setBorderBottom(border);
        nowStyle.setBorderLeft(border);
        nowStyle.setBorderRight(border);
        if (locked != null) {
            nowStyle.setLocked(locked);
        }
        nowStyle.setWrapText(true);
        return nowStyle;
    }


}
