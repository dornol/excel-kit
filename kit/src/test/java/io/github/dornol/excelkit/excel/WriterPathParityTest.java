package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.time.LocalDateTime;
import java.util.List;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.assertEquals;

/** The same column written through ExcelWriter and through ExcelWorkbook's sheet writer must look the same. */
class WriterPathParityTest {

    private static final LocalDateTime WHOLE = LocalDateTime.of(2024, 1, 1, 1, 1, 1);
    private static final LocalDateTime WITH_NANOS = LocalDateTime.of(2024, 1, 1, 1, 1, 1, 500_000_000);

    private static String format(byte[] xlsx, int row) throws Exception {
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(xlsx))) {
            Cell cell = wb.getSheetAt(0).getRow(row).getCell(0);
            return cell.getCellStyle().getDataFormatString();
        }
    }

    private static byte[] viaWriter(Function<ColumnConfig<LocalDateTime>, ?> config) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<LocalDateTime>create().column("T", t -> t, config::apply)
                .write(List.of(WHOLE, WITH_NANOS)).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] viaWorkbook(Function<ColumnConfig<LocalDateTime>, ?> config) {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = ExcelWorkbook.create()) {
            wb.<LocalDateTime>sheet("S").column("T", t -> t, config::apply).write(List.of(WHOLE, WITH_NANOS));
            wb.finish().writeTo(out);
        }
        return out.toByteArray();
    }

    @Test
    void explicitDateTimeType_keepsTheMillisecondFormatForFractionalValuesOnBothPaths() throws Exception {
        byte[] writer = viaWriter(c -> c.type(ExcelDataType.DATETIME));
        byte[] workbook = viaWorkbook(c -> c.type(ExcelDataType.DATETIME));

        assertEquals(format(workbook, 1), format(writer, 1), "whole seconds");
        assertEquals(format(workbook, 2), format(writer, 2), "fractional seconds");
    }

    @Test
    void inferredDateTime_keepsTheMillisecondFormatForFractionalValuesOnBothPaths() throws Exception {
        byte[] writer = viaWriter(c -> c);
        byte[] workbook = viaWorkbook(c -> c);

        assertEquals(format(workbook, 1), format(writer, 1));
        assertEquals(format(workbook, 2), format(writer, 2));
    }
}
