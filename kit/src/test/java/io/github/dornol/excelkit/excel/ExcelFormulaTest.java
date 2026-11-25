package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.feature.ExcelFormula;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.*;

/** A formula with its result: what a viewer that does not calculate shows, and what a reader of the file gets. */
class ExcelFormulaTest {

    private static XSSFCell write(Function<Object, Object> value) throws Exception {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<Object>create().column("F", value::apply).write(List.of(new Object())).writeTo(out);
        var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
        return wb.getSheetAt(0).getRow(1).getCell(0);
    }

    @Nested
    class Value {
        @Test
        void aLeadingEqualsSignAndSpacesAreDropped() {
            assertEquals("E2*2", ExcelFormula.of(" = E2*2 ").formula());
        }

        @Test
        void aBlankFormulaAndAResultThatIsNotAFiniteNumberAreRefused() {
            assertThrows(IllegalArgumentException.class, () -> ExcelFormula.of(" "));
            assertThrows(IllegalArgumentException.class, () -> ExcelFormula.of("="));
            assertThrows(IllegalArgumentException.class, () -> ExcelFormula.of("A1", Double.NaN));
            assertThrows(IllegalArgumentException.class, () -> ExcelFormula.of("A1", Double.POSITIVE_INFINITY));
            assertThrows(IllegalArgumentException.class, () -> ExcelFormula.of("A1", "x".repeat(32_768)));
            assertThrows(NullPointerException.class, () -> ExcelFormula.of(null));
            assertThrows(NullPointerException.class, () -> ExcelFormula.of("A1", (String) null));
            assertThrows(NullPointerException.class, () -> ExcelFormula.of("A1", (Number) null));
        }

        @Test
        void theResultIsADoubleATextOrABoolean() {
            assertEquals(2.0, ExcelFormula.of("1+1", 2).result());
            assertEquals(new BigDecimal("2.5").doubleValue(), ExcelFormula.of("1+1.5", new BigDecimal("2.5")).result());
            assertEquals("ab", ExcelFormula.of("A1&B1", "ab").result());
            assertEquals(true, ExcelFormula.of("1=1", true).result());
            assertNull(ExcelFormula.of("1+1").result());
        }

        @Test
        void toStringIsTheResultOrTheFormulaWhenThereIsNone() {
            assertEquals("2.0", ExcelFormula.of("1+1", 2).toString());
            assertEquals("ab", ExcelFormula.of("A1&B1", "ab").toString());
            assertEquals("1+1", ExcelFormula.of("1+1").toString());
        }

        @Test
        void equalsAndHashCodeFollowTheFormulaAndTheResult() {
            assertEquals(ExcelFormula.of("A1", 1), ExcelFormula.of("=A1", 1));
            assertEquals(ExcelFormula.of("A1", 1).hashCode(), ExcelFormula.of("A1", 1).hashCode());
            assertNotEquals(ExcelFormula.of("A1", 1), ExcelFormula.of("A1", 2));
            assertNotEquals(ExcelFormula.of("A1", 1), ExcelFormula.of("A1"));
        }
    }

    @Nested
    class InTheFile {
        @Test
        void aNumericResultIsStoredNextToTheFormula() throws Exception {
            XSSFCell cell = write(x -> ExcelFormula.of("E2*2", 2469134));
            assertEquals(CellType.FORMULA, cell.getCellType());
            assertEquals("E2*2", cell.getCellFormula());
            assertEquals(CellType.NUMERIC, cell.getCachedFormulaResultType());
            assertEquals(2469134.0, cell.getNumericCellValue());
        }

        @Test
        void aTextAndABooleanResultAreStoredToo() throws Exception {
            XSSFCell text = write(x -> ExcelFormula.of("\"a\"&\"b\"", "ab"));
            assertEquals(CellType.STRING, text.getCachedFormulaResultType());
            assertEquals("ab", text.getStringCellValue());
            XSSFCell flag = write(x -> ExcelFormula.of("1=1", true));
            assertEquals(CellType.BOOLEAN, flag.getCachedFormulaResultType());
            assertTrue(flag.getBooleanCellValue());
        }

        @Test
        void withoutAResultTheFormulaIsWrittenAsAStringFormulaIs() throws Exception {
            XSSFCell cell = write(x -> ExcelFormula.of("E2*2"));
            assertEquals(CellType.FORMULA, cell.getCellType());
            assertEquals("E2*2", cell.getCellFormula());
            assertEquals(CellType.STRING, cell.getCachedFormulaResultType());
            assertEquals("", cell.getStringCellValue(), "no result in the file, as before");
        }

        @Test
        void anEmptyTextAndATextThatNeedsEscapingAreNoStoredResult() throws Exception {
            for (String text : new String[]{"", "a\u0001b", "_x0041_"}) {
                XSSFCell cell = write(x -> ExcelFormula.of("A1", text));
                assertEquals("A1", cell.getCellFormula(), text);
                assertEquals("", cell.getStringCellValue(), "no result for " + text.length() + " chars");
            }
        }

        @Test
        void theTypeIsInferredAndAPlainStringFormulaStillWorks() throws Exception {
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Object>create().column("Typed", x -> "E2*2", c -> c.type(ExcelDataType.FORMULA))
                    .column("Inferred", x -> ExcelFormula.of("E2*2", 6)).write(List.of(new Object())).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var row = wb.getSheetAt(0).getRow(1);
                assertEquals("E2*2", row.getCell(0).getCellFormula());
                assertEquals("E2*2", row.getCell(1).getCellFormula());
                assertEquals(6.0, row.getCell(1).getNumericCellValue());
            }
        }

        @Test
        void aFormulaWithAResultOnRolledOverSheetsKeepsEachResult() throws Exception {
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Integer>create().maxRows(2).column("F", n -> ExcelFormula.of("A1*2", n * 2))
                    .write(List.of(1, 2, 3)).writeTo(out);
            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertEquals(2, wb.getNumberOfSheets());
                assertEquals(6.0, wb.getSheetAt(1).getRow(1).getCell(0).getNumericCellValue());
            }
        }

        @Test
        void aReaderOfTheFileGetsTheResult() throws Exception {
            var out = new ByteArrayOutputStream();
            ExcelWriter.<Object>create().column("F", x -> ExcelFormula.of("E2*2", 2469134)).write(List.of(new Object())).writeTo(out);
            List<Map<String, String>> rows = new ArrayList<>();
            ExcelReader.forMap().read(new ByteArrayInputStream(out.toByteArray()), r -> rows.add(r.data()));
            assertEquals("2469134", rows.get(0).get("F"));
        }
    }

    @Test
    void aCsvFileHoldsTheResultOrTheFormulaWhenThereIsNone() throws Exception {
        var out = new ByteArrayOutputStream();
        CsvWriter.<Object>create().column("R", x -> ExcelFormula.of("E2*2", 7)).column("P", x -> ExcelFormula.of("E2"))
                .write(List.of(new Object())).writeTo(out);
        String csv = out.toString(StandardCharsets.UTF_8);
        assertTrue(csv.contains("7.0"), csv);
        assertTrue(csv.contains("'E2") || csv.contains("E2"), csv);
        assertFalse(csv.contains("E2*2"), "the result, not the formula: " + csv);
    }
}
