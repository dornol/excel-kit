package io.github.dornol.excelkit.excel;

/**
 * Represents an RGB color for Excel header backgrounds and cell styling.
 * <p>
 * Provides commonly used color presets as static constants, and supports
 * custom RGB values via {@link #of(int, int, int)}.
 *
 * <pre>{@code
 * // Using a preset
 * ExcelWriter.create().headerColor(ExcelColor.STEEL_BLUE);
 *
 * // Using a custom RGB color
 * ExcelWriter.create().headerColor(ExcelColor.of(180, 200, 220));
 * }</pre>
 *
 * @author dhkim
 */
public final class ExcelColor {

    /** White (255, 255, 255). */
    public static final ExcelColor WHITE = new ExcelColor(255, 255, 255);
    /** Black (0, 0, 0). */
    public static final ExcelColor BLACK = new ExcelColor(0, 0, 0);

    /** Light gray (217, 217, 217). */
    public static final ExcelColor LIGHT_GRAY = new ExcelColor(217, 217, 217);
    /** Gray (128, 128, 128). */
    public static final ExcelColor GRAY = new ExcelColor(128, 128, 128);
    /** Dark gray (64, 64, 64). */
    public static final ExcelColor DARK_GRAY = new ExcelColor(64, 64, 64);

    /** Red (255, 0, 0). */
    public static final ExcelColor RED = new ExcelColor(255, 0, 0);
    /** Green (0, 128, 0). */
    public static final ExcelColor GREEN = new ExcelColor(0, 128, 0);
    /** Blue (0, 0, 255). */
    public static final ExcelColor BLUE = new ExcelColor(0, 0, 255);
    /** Yellow (255, 255, 0). */
    public static final ExcelColor YELLOW = new ExcelColor(255, 255, 0);
    /** Orange (255, 165, 0). */
    public static final ExcelColor ORANGE = new ExcelColor(255, 165, 0);

    /** Light red for backgrounds. */
    public static final ExcelColor LIGHT_RED = new ExcelColor(255, 199, 206);
    /** Light green for backgrounds. */
    public static final ExcelColor LIGHT_GREEN = new ExcelColor(198, 239, 206);
    /** Light blue for backgrounds. */
    public static final ExcelColor LIGHT_BLUE = new ExcelColor(189, 215, 238);
    /** Light yellow for backgrounds. */
    public static final ExcelColor LIGHT_YELLOW = new ExcelColor(255, 235, 156);
    /** Light orange for backgrounds. */
    public static final ExcelColor LIGHT_ORANGE = new ExcelColor(252, 228, 214);
    /** Light purple for backgrounds. */
    public static final ExcelColor LIGHT_PURPLE = new ExcelColor(228, 210, 245);

    /** Purple (128, 0, 128). */
    public static final ExcelColor PURPLE = new ExcelColor(128, 0, 128);
    /** Pink (255, 192, 203). */
    public static final ExcelColor PINK = new ExcelColor(255, 192, 203);
    /** Teal (0, 128, 128). */
    public static final ExcelColor TEAL = new ExcelColor(0, 128, 128);
    /** Navy (0, 0, 128). */
    public static final ExcelColor NAVY = new ExcelColor(0, 0, 128);

    /** Coral (255, 127, 80). */
    public static final ExcelColor CORAL = new ExcelColor(255, 127, 80);
    /** Steel blue (70, 130, 180). */
    public static final ExcelColor STEEL_BLUE = new ExcelColor(70, 130, 180);
    /** Forest green (34, 139, 34). */
    public static final ExcelColor FOREST_GREEN = new ExcelColor(34, 139, 34);
    /** Gold (255, 215, 0). */
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

    /** Returns the red component (0-255).
     * @return red value */
    public int getR() {
        return r;
    }

    /** Returns the green component (0-255).
     * @return green value */
    public int getG() {
        return g;
    }

    /** Returns the blue component (0-255).
     * @return blue value */
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
