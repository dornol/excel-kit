package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.jspecify.annotations.Nullable;

/**
 * Bundles all cell styling parameters into a single record to avoid
 * bloating method signatures in {@link ExcelStyleSupporter}.
 *
 * @author dhkim
 * @since 0.7.0
 */
record CellStyleParams(
        HorizontalAlignment alignment,
        @Nullable String format,
        int @Nullable [] backgroundColor,
        @Nullable Boolean bold,
        @Nullable Integer fontSize,
        @Nullable ExcelBorderStyle borderStyle,
        @Nullable Boolean locked,
        @Nullable Short rotation,
        @Nullable ExcelBorderStyle borderTop,
        @Nullable ExcelBorderStyle borderBottom,
        @Nullable ExcelBorderStyle borderLeft,
        @Nullable ExcelBorderStyle borderRight,
        int @Nullable [] fontColor,
        @Nullable Boolean strikethrough,
        @Nullable Boolean underline,
        @Nullable VerticalAlignment verticalAlignment,
        @Nullable Boolean wrapText,
        @Nullable String fontName,
        @Nullable Short indentation
) {
    /** Creates params with only alignment and format; all styling fields are null/default. */
    static CellStyleParams of(HorizontalAlignment alignment, @Nullable String format) {
        return new CellStyleParams(alignment, format,
                null, null, null, null, null, null, null, null, null, null,
                null, null, null, null, null, null, null);
    }

    /** Creates params with core styling fields; extended fields (rotation, per-side borders, font, etc.) are null. */
    static CellStyleParams of(HorizontalAlignment alignment, @Nullable String format,
                               int @Nullable [] backgroundColor, @Nullable Boolean bold,
                               @Nullable Integer fontSize, @Nullable ExcelBorderStyle borderStyle,
                               @Nullable Boolean locked) {
        return new CellStyleParams(alignment, format, backgroundColor, bold, fontSize,
                borderStyle, locked, null, null, null, null, null, null, null, null,
                null, null, null, null);
    }
}
