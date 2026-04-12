package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for FORMULA type behavior.
 * <p>
 * Note: POI accepts DDE formulas (e.g., cmd|'/c calc') without validation.
 * This is documented in ExcelDataType.FORMULA javadoc as a security warning.
 * DDE is intentionally allowed because legitimate use cases exist (e.g., Bloomberg DDE links).
 */
class FormulaInjectionTest {

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
    void pipeFormula_isAcceptedByPoi() throws Exception {
        // POI accepts DDE formulas — this is by design (Bloomberg, Reuters use cases).
        // Security is the developer's responsibility per javadoc warning.
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter.<String>builder().build()
                .column("Calc", s -> s, c -> c.type(ExcelDataType.FORMULA))
                .write(Stream.of("cmd|'/c calc'"))
                .write(bos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bos.toByteArray()))) {
            var cell = wb.getSheetAt(0).getRow(1).getCell(0);
            assertEquals(CellType.FORMULA, cell.getCellType(),
                    "POI accepts DDE formulas — developer must validate input per javadoc");
        }
    }
}
