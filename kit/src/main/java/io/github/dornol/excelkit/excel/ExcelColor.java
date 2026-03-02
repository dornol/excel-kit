package io.github.dornol.excelkit.excel;

/**
 * Predefined color presets for Excel header backgrounds and cell styling.
 * <p>
 * Provides commonly used colors so that users do not need to look up RGB values manually.
 * Can be used with {@link ExcelWriter} constructors and
 * {@link ExcelColumn.ExcelColumnBuilder#backgroundColor(ExcelColor)}.
 *
 * @author dhkim
 */
public enum ExcelColor {

    WHITE(255, 255, 255),
    BLACK(0, 0, 0),

    // Gray
    LIGHT_GRAY(217, 217, 217),
    GRAY(128, 128, 128),
    DARK_GRAY(64, 64, 64),

    // Basic colors
    RED(255, 0, 0),
    GREEN(0, 128, 0),
    BLUE(0, 0, 255),
    YELLOW(255, 255, 0),
    ORANGE(255, 165, 0),

    // Light colors (for backgrounds)
    LIGHT_RED(255, 199, 206),
    LIGHT_GREEN(198, 239, 206),
    LIGHT_BLUE(189, 215, 238),
    LIGHT_YELLOW(255, 235, 156),
    LIGHT_ORANGE(252, 228, 214),
    LIGHT_PURPLE(228, 210, 245),

    // Commonly used colors
    CORAL(255, 127, 80),
    STEEL_BLUE(70, 130, 180),
    FOREST_GREEN(34, 139, 34),
    GOLD(255, 215, 0),
    ;

    private final int r;
    private final int g;
    private final int b;

    ExcelColor(int r, int g, int b) {
        this.r = r;
        this.g = g;
        this.b = b;
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
