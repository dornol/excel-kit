package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.HorizontalAlignment;
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
        @Nullable Boolean underline
) {
}
