package io.github.dornol.excelkit.excel.read;

import org.junit.jupiter.api.Test;

import java.util.*;

import static org.junit.jupiter.api.Assertions.*;

class ExcelSheetInfoTest {

    // ============================================================
    // ExcelSheetInfo record tests
    // ============================================================
    @Test
    void excelSheetInfo_record() {
        var info = new ExcelSheetInfo(2, "Sheet3");
        assertEquals(2, info.index());
        assertEquals("Sheet3", info.name());
    }
}
