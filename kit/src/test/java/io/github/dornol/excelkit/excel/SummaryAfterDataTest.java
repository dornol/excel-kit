package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.*;

import java.io.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("Bugfix: summary + afterData row overlap")
class SummaryAfterDataTest {

    @Test
    void summary_afterData_shouldNotOverlapRows() throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<int[]>create()
                .column("Name", arr -> "Item" + arr[0])
                .column("Amount", arr -> arr[0], c -> c.type(ExcelDataType.INTEGER))
                .afterData((ctx) -> {
                    // Write a custom row after data
                    var row = ctx.getSheet().createRow(ctx.getCurrentRow());
                    row.createCell(0).setCellValue("Custom Footer");
                    return ctx.getCurrentRow() + 1;
                })
                .summary(s -> s.sum("Amount"))
                .write(Stream.of(new int[]{10}, new int[]{20}, new int[]{30}))
                .writeTo(bos);

        // Read back and verify no overlapping rows
        try (var wb = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(bos.toByteArray()))) {
            var sheet = wb.getSheetAt(0);
            // Row 0: header, Row 1-3: data, Row 4: afterData footer, Row 5: summary
            String footerValue = sheet.getRow(4).getCell(0).getStringCellValue();
            assertEquals("Custom Footer", footerValue, "afterData row should be at row 4");

            // Summary row should be AFTER the footer, not overlapping
            String summaryLabel = sheet.getRow(5).getCell(0).getStringCellValue();
            assertNotNull(summaryLabel, "Summary row should exist at row 5");
        }
    }
}
