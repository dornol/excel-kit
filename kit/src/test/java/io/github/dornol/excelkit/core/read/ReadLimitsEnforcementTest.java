package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

class ReadLimitsEnforcementTest {

    @Test void totalCellLimitAbortsCsvRead() {
        assertThrows(ReadLimitExceededException.class, () ->
                CsvReader.<String>mapping(row -> row.get("Name").asString())
                        .limits(new ReadLimits(-1, -1, -1, -1, 1))
                        .read(csv("Name\nA\nB\n"), result -> {}));
    }

    @Test void columnAndCellLimitsAbortCsvRead() {
        assertThrows(ExcelKitException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(-1, -1, 1, -1))
                .read(csv("A,B\n1,2\n"), result -> fail("must fail before rows")));
        assertThrows(ReadLimitExceededException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(-1, -1, -1, 2))
                .read(csv("A\nlong\n"), result -> fail("must not deliver an oversized cell")));
    }

    @Test void columnLimitAppliesToDataRowsWiderThanHeader() {
        assertThrows(ReadLimitExceededException.class, () -> CsvReader.forMap()
                .limits(new ReadLimits(-1, -1, 1, -1))
                .read(csv("A\n1,2\n"), result -> fail("must reject a wide data row")));
    }

    @Test void excelSheetLimitAbortsBeforeRows() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            workbook.createSheet("One").createRow(0).createCell(0).setCellValue("A");
            workbook.createSheet("Two").createRow(0).createCell(0).setCellValue("A");
            workbook.write(output);
        }
        assertThrows(ReadLimitExceededException.class, () ->
                ExcelReader.forMap().limits(new ReadLimits(-1, 1, -1, -1))
                        .read(new ByteArrayInputStream(output.toByteArray()), result -> fail()));
    }

    private static ByteArrayInputStream csv(String value) {
        return new ByteArrayInputStream(value.getBytes(StandardCharsets.UTF_8));
    }
}
