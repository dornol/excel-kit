package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Settings are captured when write() starts; changing the writer while it writes has no effect. */
class SheetConfigSnapshotTest {

    @Test
    void maxRowsAndTableChangedWhileWriting_areIgnored() throws Exception {
        ExcelWriter<Integer> writer = ExcelWriter.<Integer>create().maxRows(10);
        writer.column("N", n -> {
            writer.maxRows(1);
            writer.table("Late");
            return n;
        });
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(List.of(1, 2, 3)).writeTo(out);

        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(1, wb.getNumberOfSheets());
            XSSFSheet sheet = wb.getSheetAt(0);
            assertEquals(3, sheet.getLastRowNum());
            assertTrue(sheet.getTables().isEmpty());
        }
    }
}
