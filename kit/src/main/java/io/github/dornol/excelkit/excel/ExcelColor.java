package io.github.dornol.excelkit.excel;

/**
 * Represents an RGB color for Excel header backgrounds and cell styling.
 * <p>
 * Provides commonly used color presets as static constants, and supports
 * custom RGB values via {@link #of(int, int, int)}.
 *
 * <pre>{@code
 * // Using a preset
 * new ExcelWriter<>(ExcelColor.STEEL_BLUE);
 *
 * // Using a custom RGB color
 * new ExcelWriter<>(ExcelColor.of(180, 200, 220));
 * }</pre>
 *
 * @author dhkim
 */
public final class ExcelColor {

    // White / Black
    public static final ExcelColor WHITE = new ExcelColor(255, 255, 255);
    public static final ExcelColor BLACK = new ExcelColor(0, 0, 0);

    // Gray
    public static final ExcelColor LIGHT_GRAY = new ExcelColor(217, 217, 217);
    public static final ExcelColor GRAY = new ExcelColor(128, 128, 128);
    public static final ExcelColor DARK_GRAY = new ExcelColor(64, 64, 64);

    // Basic colors
    public static final ExcelColor RED = new ExcelColor(255, 0, 0);
    public static final ExcelColor GREEN = new ExcelColor(0, 128, 0);
    public static final ExcelColor BLUE = new ExcelColor(0, 0, 255);
    public static final ExcelColor YELLOW = new ExcelColor(255, 255, 0);
    public static final ExcelColor ORANGE = new ExcelColor(255, 165, 0);

    // Light colors (for backgrounds)
    public static final ExcelColor LIGHT_RED = new ExcelColor(255, 199, 206);
    public static final ExcelColor LIGHT_GREEN = new ExcelColor(198, 239, 206);
    public static final ExcelColor LIGHT_BLUE = new ExcelColor(189, 215, 238);
    public static final ExcelColor LIGHT_YELLOW = new ExcelColor(255, 235, 156);
    public static final ExcelColor LIGHT_ORANGE = new ExcelColor(252, 228, 214);
    public static final ExcelColor LIGHT_PURPLE = new ExcelColor(228, 210, 245);

    // Additional basic colors
    public static final ExcelColor PURPLE = new ExcelColor(128, 0, 128);
    public static final ExcelColor PINK = new ExcelColor(255, 192, 203);
    public static final ExcelColor TEAL = new ExcelColor(0, 128, 128);
    public static final ExcelColor NAVY = new ExcelColor(0, 0, 128);

    // Commonly used colors
    public static final ExcelColor CORAL = new ExcelColor(255, 127, 80);
    public static final ExcelColor STEEL_BLUE = new ExcelColor(70, 130, 180);
    public static final ExcelColor FOREST_GREEN = new ExcelColor(34, 139, 34);
    public static final ExcelColor GOLD = new ExcelColor(255, 215, 0);

    private final int r;
    private final int g;
    private final int b;

    private ExcelColor(int r, int g, int b) {
        if (r < 0 || r > 255 || g < 0 || g > 255 || b < 0 || b > 255) {
            throw new IllegalArgumentException(
                    "RGB values must be between 0 and 255, got (" + r + ", " + g + ", " + b + ")");
        }
        this.r = r;
        this.g = g;
        this.b = b;
    }

    /**
     * Creates a custom color from RGB values.
     *
     * @param r Red component (0–255)
     * @param g Green component (0–255)
     * @param b Blue component (0–255)
     * @return a new ExcelColor instance
     */
    public static ExcelColor of(int r, int g, int b) {
        return new ExcelColor(r, g, b);
    }

    public int getR() {
        return r;
    }

    public int getG() {
        return g;
    }

    public int getB() {
        return b;
    }

    /**
     * Returns the color as an RGB array.
     *
     * @return {@code {r, g, b}} array
     */
    public int[] toRgb() {
        return new int[]{r, g, b};
    }
}
