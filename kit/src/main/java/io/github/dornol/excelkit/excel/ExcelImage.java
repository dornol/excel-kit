package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Workbook;

/**
 * Represents image data to be embedded in an Excel cell.
 *
 * @param data      the image bytes (PNG, JPEG, etc.)
 * @param imageType the image type constant from {@link Workbook}
 *                  (e.g., {@link Workbook#PICTURE_TYPE_PNG}, {@link Workbook#PICTURE_TYPE_JPEG})
 * @author dhkim
 * @since 0.6.0
 */
public record ExcelImage(byte[] data, int imageType) {

    /**
     * Creates a PNG image.
     *
     * @param data the PNG image bytes
     * @return an ExcelImage with PNG type
     */
    public static ExcelImage png(byte[] data) {
        return new ExcelImage(data, Workbook.PICTURE_TYPE_PNG);
    }

    /**
     * Creates a JPEG image.
     *
     * @param data the JPEG image bytes
     * @return an ExcelImage with JPEG type
     */
    public static ExcelImage jpeg(byte[] data) {
        return new ExcelImage(data, Workbook.PICTURE_TYPE_JPEG);
    }
}
