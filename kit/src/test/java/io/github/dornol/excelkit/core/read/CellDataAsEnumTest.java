package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.excel.ExcelDataType;
import org.apache.poi.ss.usermodel.*;
import org.junit.jupiter.api.Test;

import java.io.*;

import static org.junit.jupiter.api.Assertions.*;

class CellDataAsEnumTest {

    // ========================================================================
    // CellData.asEnum()
    // ========================================================================
    @Test
    void asEnum_shouldParseByName() {
        CellData cell = new CellData(0, "ACTIVE");
        assertEquals(Status.ACTIVE, cell.asEnum(Status.class));
    }

    @Test
    void asEnum_shouldBeCaseInsensitive() {
        assertEquals(Status.ACTIVE, new CellData(0, "active").asEnum(Status.class));
        assertEquals(Status.INACTIVE, new CellData(0, "Inactive").asEnum(Status.class));
        assertEquals(Status.PENDING, new CellData(0, "PENDING").asEnum(Status.class));
    }

    @Test
    void asEnum_blank_shouldReturnNull() {
        assertNull(new CellData(0, "").asEnum(Status.class));
        assertNull(new CellData(0, "  ").asEnum(Status.class));
    }

    @Test
    void asEnum_invalidValue_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new CellData(0, "UNKNOWN").asEnum(Status.class));
    }

    @Test
    void asEnum_withTrimming() {
        assertEquals(Status.ACTIVE, new CellData(0, " ACTIVE ").asEnum(Status.class));
    }

    // -----------------------------------------------------------------------
    // Additional edge-case coverage
    // -----------------------------------------------------------------------

    @Test
    void cellData_asEnum() {
        CellData cell = new CellData(0, "STRING");
        ExcelDataType result = cell.asEnum(ExcelDataType.class);
        assertEquals(ExcelDataType.STRING, result);
    }

    @Test
    void cellData_asEnum_caseInsensitive() {
        CellData cell = new CellData(0, "string");
        ExcelDataType result = cell.asEnum(ExcelDataType.class);
        assertEquals(ExcelDataType.STRING, result);
    }

    @Test
    void cellData_asEnum_blankReturnsNull() {
        assertNull(new CellData(0, "").asEnum(ExcelDataType.class));
    }

    @Test
    void cellData_asEnum_invalidThrows() {
        CellData cell = new CellData(0, "INVALID_VALUE");
        assertThrows(IllegalArgumentException.class, () -> cell.asEnum(ExcelDataType.class));
    }

    public enum Status { ACTIVE, INACTIVE, PENDING }
}
