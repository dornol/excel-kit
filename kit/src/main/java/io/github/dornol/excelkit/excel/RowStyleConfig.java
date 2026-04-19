package io.github.dornol.excelkit.excel;

import org.jspecify.annotations.Nullable;

/**
 * Configuration for conditional row-level styling.
 * <p>
 * Unlike {@code rowColor()} which only sets background color, {@code RowStyleConfig} supports
 * bold, font size, font color, and background color — applied to every cell in the row
 * when the condition matches.
 *
 * <pre>{@code
 * writer.rowStyle(
 *     product -> product.price() > 10000,
 *     style -> style.bold(true).backgroundColor(ExcelColor.LIGHT_YELLOW)
 * );
 * }</pre>
 *
 * @author dhkim
 */
public class RowStyleConfig {

    @Nullable Boolean bold;
    @Nullable Float fontSize;
    @Nullable ExcelColor fontColor;
    @Nullable ExcelColor backgroundColor;
    @Nullable Boolean italic;
    @Nullable Boolean strikethrough;

    /**
     * Sets bold font for the row.
     */
    public RowStyleConfig bold(boolean bold) {
        this.bold = bold;
        return this;
    }

    /**
     * Sets the font size for the row.
     */
    public RowStyleConfig fontSize(float fontSize) {
        this.fontSize = fontSize;
        return this;
    }

    /**
     * Sets the font color for the row.
     */
    public RowStyleConfig fontColor(ExcelColor color) {
        this.fontColor = color;
        return this;
    }

    /**
     * Sets the font color for the row using RGB values.
     */
    public RowStyleConfig fontColor(int r, int g, int b) {
        this.fontColor = ExcelColor.of(r, g, b);
        return this;
    }

    /**
     * Sets the background color for the row.
     */
    public RowStyleConfig backgroundColor(ExcelColor color) {
        this.backgroundColor = color;
        return this;
    }

    /**
     * Sets the background color for the row using RGB values.
     */
    public RowStyleConfig backgroundColor(int r, int g, int b) {
        this.backgroundColor = ExcelColor.of(r, g, b);
        return this;
    }

    /**
     * Sets italic font for the row.
     */
    public RowStyleConfig italic(boolean italic) {
        this.italic = italic;
        return this;
    }

    /**
     * Sets strikethrough font for the row.
     */
    public RowStyleConfig strikethrough(boolean strikethrough) {
        this.strikethrough = strikethrough;
        return this;
    }

    boolean hasAnyStyle() {
        return bold != null || fontSize != null || fontColor != null
                || backgroundColor != null || italic != null || strikethrough != null;
    }

    /**
     * Builds a cache key fragment representing this style configuration.
     */
    String cacheKey() {
        StringBuilder sb = new StringBuilder("rs_");
        if (bold != null) sb.append("b").append(bold ? "1" : "0");
        if (fontSize != null) sb.append("f").append(fontSize.intValue());
        if (fontColor != null) sb.append("fc").append(fontColor.getR()).append("_").append(fontColor.getG()).append("_").append(fontColor.getB());
        if (backgroundColor != null) sb.append("bg").append(backgroundColor.getR()).append("_").append(backgroundColor.getG()).append("_").append(backgroundColor.getB());
        if (italic != null) sb.append("i").append(italic ? "1" : "0");
        if (strikethrough != null) sb.append("s").append(strikethrough ? "1" : "0");
        return sb.toString();
    }
}
