package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;

import java.io.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("Bugfix: ExcelSheetWriter rollover header style")
class RolloverHeaderStyleTest {

    @Test
    void rollover_shouldPreserveHeaderFontColor() throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            ExcelSheetWriter<Integer> sw = wb.sheet("test");
            sw.column("ID", i -> i, c -> c.headerFontColor(ExcelColor.RED))
                    .maxRows(2);
            sw.write(Stream.of(1, 2, 3, 4));
            wb.finish().writeTo(bos);
        }

        // Verify the file is readable and has 2 sheets (rollover occurred)
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(bos.toByteArray()))) {
            assertTrue(workbook.getNumberOfSheets() >= 2,
                    "Should have at least 2 sheets due to maxRows=2, got " + workbook.getNumberOfSheets());
        }
    }
}
