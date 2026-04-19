package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Workbook;

import java.io.IOException;
import java.io.InputStream;
import java.io.UncheckedIOException;
import java.net.URI;

/**
 * Represents image data to be embedded in an Excel cell.
 * <p>
 * Use {@link #png(byte[])}, {@link #jpeg(byte[])}, or {@link #fromUrl(String)} to create instances.
 * Optionally call {@link #size(int, int)} to control the image span in cells.
 *
 * <pre>{@code
 * // Fixed size
 * ExcelImage.png(bytes).size(2, 3)   // 2 columns wide, 3 rows tall
 *
 * // From URL (auto-detects PNG/JPEG)
 * ExcelImage.fromUrl("https://example.com/photo.png")
 * }</pre>
 *
 * @author dhkim
 * @since 0.6.0
 */
public class ExcelImage {

    private final byte[] data;
    private final int imageType;
    private int width = 1;
    private int height = 1;

    /**
     * Creates an ExcelImage with the given data and type.
     *
     * @param data      the image bytes (PNG, JPEG, etc.)
     * @param imageType the image type constant from {@link Workbook}
     */
    public ExcelImage(byte[] data, int imageType) {
        if (data == null || data.length == 0) {
            throw new IllegalArgumentException("Image data must not be null or empty");
        }
        this.data = data;
        this.imageType = imageType;
    }

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

    /**
     * Downloads an image from a URL and creates an ExcelImage.
     * <p>
     * The image type is auto-detected from the URL extension (defaults to PNG).
     * Uses a 10-second connect/read timeout.
     *
     * <pre>{@code
     * writer.column("Photo", user -> ExcelImage.fromUrl(user.getPhotoUrl()),
     *     c -> c.type(ExcelDataType.IMAGE));
     * }</pre>
     *
     * @param url the image URL
     * @return an ExcelImage with the downloaded data
     * @throws UncheckedIOException if the download fails
     */
    public static ExcelImage fromUrl(String url) {
        try {
            var conn = URI.create(url).toURL().openConnection();
            conn.setConnectTimeout(10_000);
            conn.setReadTimeout(10_000);
            byte[] data;
            try (InputStream is = conn.getInputStream()) {
                data = is.readAllBytes();
            }
            int type = detectImageType(url, data);
            return new ExcelImage(data, type);
        } catch (IOException e) {
            throw new UncheckedIOException("Failed to download image from: " + url, e);
        }
    }

    /**
     * Sets the image span in cells (columns wide x rows tall).
     * Defaults to 1x1.
     *
     * @param width  number of columns the image spans
     * @param height number of rows the image spans
     * @return this image for chaining
     */
    public ExcelImage size(int width, int height) {
        if (width < 1) throw new IllegalArgumentException("width must be >= 1");
        if (height < 1) throw new IllegalArgumentException("height must be >= 1");
        this.width = width;
        this.height = height;
        return this;
    }

    /**
     * Returns the image bytes.
     */
    public byte[] data() {
        return data;
    }

    /**
     * Returns the image type constant.
     */
    public int imageType() {
        return imageType;
    }

    /**
     * Returns the number of columns the image spans. Default is 1.
     */
    public int width() {
        return width;
    }

    /**
     * Returns the number of rows the image spans. Default is 1.
     */
    public int height() {
        return height;
    }

    private static int detectImageType(String url, byte[] data) {
        String lower = url.toLowerCase(java.util.Locale.ROOT);
        if (lower.endsWith(".jpg") || lower.endsWith(".jpeg")) {
            return Workbook.PICTURE_TYPE_JPEG;
        }
        if (lower.endsWith(".png")) {
            return Workbook.PICTURE_TYPE_PNG;
        }
        // Detect from magic bytes
        if (data.length >= 2 && (data[0] & 0xFF) == 0xFF && (data[1] & 0xFF) == 0xD8) {
            return Workbook.PICTURE_TYPE_JPEG;
        }
        // Default to PNG
        return Workbook.PICTURE_TYPE_PNG;
    }
}
