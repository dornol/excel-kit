package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Predefined border styles for Excel cells.
 * <p>
 * Wraps Apache POI's {@link BorderStyle} to provide a simplified API
 * for configuring cell borders in {@link ExcelColumn.ExcelColumnBuilder}
 * and {@link ExcelSheetWriter.ColumnConfig}.
 *
 * @author dhkim
 * @since 0.6.0
 */
public enum ExcelBorderStyle {

    NONE(BorderStyle.NONE),
    THIN(BorderStyle.THIN),
    MEDIUM(BorderStyle.MEDIUM),
    THICK(BorderStyle.THICK),
    DASHED(BorderStyle.DASHED),
    DOTTED(BorderStyle.DOTTED),
    DOUBLE(BorderStyle.DOUBLE),
    HAIR(BorderStyle.HAIR),
    MEDIUM_DASHED(BorderStyle.MEDIUM_DASHED),
    DASH_DOT(BorderStyle.DASH_DOT),
    ;

    private final BorderStyle poiStyle;

    ExcelBorderStyle(BorderStyle poiStyle) {
        this.poiStyle = poiStyle;
    }

    /**
     * Returns the corresponding Apache POI {@link BorderStyle}.
     */
    BorderStyle toPoiBorderStyle() {
        return poiStyle;
    }
}
