package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.BorderStyle;

/**
 * Predefined border styles for Excel cells.
 * <p>
 * Wraps Apache POI's {@link BorderStyle} to provide a simplified API
 * for configuring cell borders in {@link ExcelColumn.ExcelColumnBuilder}
 * and {@link ColumnConfig}.
 *
 * @author dhkim
 * @since 0.6.0
 */
public enum ExcelBorderStyle {

    /** No border. */
    NONE(BorderStyle.NONE),
    /** Thin border. */
    THIN(BorderStyle.THIN),
    /** Medium border. */
    MEDIUM(BorderStyle.MEDIUM),
    /** Thick border. */
    THICK(BorderStyle.THICK),
    /** Dashed border. */
    DASHED(BorderStyle.DASHED),
    /** Dotted border. */
    DOTTED(BorderStyle.DOTTED),
    /** Double border. */
    DOUBLE(BorderStyle.DOUBLE),
    /** Hairline border. */
    HAIR(BorderStyle.HAIR),
    /** Medium dashed border. */
    MEDIUM_DASHED(BorderStyle.MEDIUM_DASHED),
    /** Dash-dot border. */
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
