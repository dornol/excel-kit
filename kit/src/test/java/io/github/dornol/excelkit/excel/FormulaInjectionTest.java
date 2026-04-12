package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for DDE formula injection defense.
 * POI does NOT reject DDE formulas (e.g., cmd|'/c calc') — our guard is required.
 */
class FormulaInjectionTest {

    @Test
    void dde_pipeFormula_shouldBeBlockedAsString() throws Exception {
        // DDE injection: cmd|'/c calc' contains pipe, our guard rejects it,
        // ExcelColumn's catch block converts to string cell
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>builder().build()
                .column("Calc", s -> s, c -> c.type(ExcelDataType.FORMULA))
                .write(Stream.of("cmd|'/c calc'"))
                .write(bos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals(CellType.STRING, cell.getCellType(),
                    "DDE formula should be rejected and stored as plain string");
            assertEquals("cmd|'/c calc'", cell.getStringCellValue());
        }
    }

    @Test
    void normalFormula_shouldWork() throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>builder().build()
                .column("Val", s -> "100")
                .column("Calc", s -> "SUM(A2:A2)", c -> c.type(ExcelDataType.FORMULA))
                .write(Stream.of("test"))
                .write(bos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(1);
            assertEquals(CellType.FORMULA, cell.getCellType());
            assertEquals("SUM(A2:A2)", cell.getCellFormula());
        }
    }

    @Test
    void formulaWithoutPipe_shouldWork() throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>builder().build()
                .column("Calc", s -> "A1+B1", c -> c.type(ExcelDataType.FORMULA))
                .write(Stream.of("test"))
                .write(bos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals(CellType.FORMULA, cell.getCellType());
        }
    }
}
