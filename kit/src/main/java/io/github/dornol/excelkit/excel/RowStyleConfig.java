package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
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
 * @since 0.16.15
 */
public class RowStyleConfig {

    @Nullable Boolean bold;
    @Nullable Float fontSize;
    @Nullable ExcelColor fontColor;
    @Nullable ExcelColor backgroundColor;
    @Nullable Boolean italic;
    @Nullable Boolean strikethrough;
    /** Set by the writer only (no public setter): a row whose cells must stay locked under sheet protection. */
    @Nullable Boolean locked;
    private @Nullable String cacheKey;

    /**
     * Sets bold font for the row.
     */
    public RowStyleConfig bold(boolean bold) {
        this.bold = bold;
        cacheKey = null;
        return this;
    }

    /**
     * Sets the font size for the row.
     */
    public RowStyleConfig fontSize(float fontSize) {
        this.fontSize = fontSize;
        cacheKey = null;
        return this;
    }

    /**
     * Sets the font color for the row.
     */
    public RowStyleConfig fontColor(@Nullable ExcelColor color) {
        this.fontColor = color;
        cacheKey = null;
        return this;
    }

    /**
     * Sets the font color for the row using RGB values.
     */
    public RowStyleConfig fontColor(int r, int g, int b) {
        this.fontColor = ExcelColor.of(r, g, b);
        cacheKey = null;
        return this;
    }

    /**
     * Sets the background color for the row.
     */
    public RowStyleConfig backgroundColor(@Nullable ExcelColor color) {
        this.backgroundColor = color;
        cacheKey = null;
        return this;
    }

    /**
     * Sets the background color for the row using RGB values.
     */
    public RowStyleConfig backgroundColor(int r, int g, int b) {
        this.backgroundColor = ExcelColor.of(r, g, b);
        cacheKey = null;
        return this;
    }

    /**
     * Sets italic font for the row.
     */
    public RowStyleConfig italic(boolean italic) {
        this.italic = italic;
        cacheKey = null;
        return this;
    }

    /**
     * Sets strikethrough font for the row.
     */
    public RowStyleConfig strikethrough(boolean strikethrough) {
        this.strikethrough = strikethrough;
        cacheKey = null;
        return this;
    }

    /** A copy of this style whose cells are locked, whatever the column says. */
    RowStyleConfig lockedCopy() {
        RowStyleConfig copy = new RowStyleConfig();
        copy.bold = bold;
        copy.fontSize = fontSize;
        copy.fontColor = fontColor;
        copy.backgroundColor = backgroundColor;
        copy.italic = italic;
        copy.strikethrough = strikethrough;
        copy.locked = true;
        return copy;
    }

    boolean hasAnyStyle() {
        return bold != null || fontSize != null || fontColor != null
                || backgroundColor != null || italic != null || strikethrough != null || locked != null;
    }

    /**
     * A cache key fragment representing this style configuration. Built once and reused, because the
     * writers ask for it on every cell; any setter drops the remembered value.
     */
    String cacheKey() {
        String key = cacheKey;
        if (key == null) {
            key = buildCacheKey();
            cacheKey = key;
        }
        return key;
    }

    private String buildCacheKey() {
        StringBuilder sb = new StringBuilder("rs_");
        if (bold != null) sb.append("b").append(bold ? "1" : "0");
        if (fontSize != null) sb.append("f").append(fontSize.intValue());
        if (fontColor != null) sb.append("fc").append(fontColor.getR()).append("_").append(fontColor.getG()).append("_").append(fontColor.getB());
        if (backgroundColor != null) sb.append("bg").append(backgroundColor.getR()).append("_").append(backgroundColor.getG()).append("_").append(backgroundColor.getB());
        if (italic != null) sb.append("i").append(italic ? "1" : "0");
        if (strikethrough != null) sb.append("s").append(strikethrough ? "1" : "0");
        if (locked != null) sb.append("l").append(locked ? "1" : "0");
        return sb.toString();
    }
}
