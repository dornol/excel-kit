package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.jspecify.annotations.Nullable;

/**
 * Configuration for the global header cell style.
 * <p>
 * Allows overriding the default header appearance (bold, alignment, border).
 * Font name, font size, and background color are set separately via
 * dedicated writer methods ({@code headerFontName}, {@code headerFontSize},
 * {@code headerColor}).
 *
 * @author dhkim
 * @since 0.17.0
 */
public class HeaderStyleConfig {

    @Nullable Boolean bold;
    @Nullable HorizontalAlignment alignment;
    @Nullable VerticalAlignment verticalAlignment;
    @Nullable ExcelBorderStyle borderStyle;
    @Nullable Boolean wrapText;

    /** Creates a new header style configuration with defaults. */
    public HeaderStyleConfig() {}

    /**
     * Sets whether header text is bold. Default: {@code true}.
     *
     * @param bold whether headers should be bold
     * @return this config for chaining
     */
    public HeaderStyleConfig bold(boolean bold) {
        this.bold = bold;
        return this;
    }

    /**
     * Sets horizontal alignment for header cells. Default: {@code CENTER}.
     *
     * @param alignment the horizontal alignment
     * @return this config for chaining
     */
    public HeaderStyleConfig alignment(HorizontalAlignment alignment) {
        this.alignment = alignment;
        return this;
    }

    /**
     * Sets vertical alignment for header cells. Default: {@code CENTER}.
     *
     * @param verticalAlignment the vertical alignment
     * @return this config for chaining
     */
    public HeaderStyleConfig verticalAlignment(VerticalAlignment verticalAlignment) {
        this.verticalAlignment = verticalAlignment;
        return this;
    }

    /**
     * Sets the border style for header cells. Default: {@code THIN}.
     *
     * @param borderStyle the border style
     * @return this config for chaining
     */
    public HeaderStyleConfig border(ExcelBorderStyle borderStyle) {
        this.borderStyle = borderStyle;
        return this;
    }

    /**
     * Sets whether header text wraps. Default: not set (Excel default).
     *
     * @param wrapText whether to wrap text
     * @return this config for chaining
     */
    public HeaderStyleConfig wrapText(boolean wrapText) {
        this.wrapText = wrapText;
        return this;
    }
}
