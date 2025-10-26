package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.streaming.SXSSFCell;

/**
 * Functional interface for setting a value into an Excel cell.
 * <p>
 * This is used internally by {@link ExcelDataType} to define how each Java type
 * should be written into an Excel cell (e.g., string, number, date).
 *
 * @author dhkim
 * @since 2025-07-19
 */
@FunctionalInterface
interface ExcelColumnSetter {

    /**
     * Sets the given value into the specified SXSSFCell.
     *
     * @param cell  The target Excel cell
     * @param value The value to write into the cell (type-specific)
     */
    void set(SXSSFCell cell, Object value);

}
