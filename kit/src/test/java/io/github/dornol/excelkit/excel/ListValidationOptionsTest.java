package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ListValidationOptionsTest {

    private record Row(String category, String item) {}

    private static XSSFWorkbook write(ExcelWriter<Row> writer) throws IOException {
        var out = new ByteArrayOutputStream();
        writer.write(Stream.of(new Row("A", "1"))).writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static ExcelWriter<Row> writer(java.util.function.Consumer<ExcelColumn.ExcelColumnBuilder<Row>> second) {
        return ExcelWriter.<Row>create()
                .column("Category", Row::category, c -> c.dropdown("A", "B"))
                .column("Item", Row::item, second);
    }

    private static Map<String, List<String>> options() {
        return Map.of("A", List.of("1", "2"), "B", List.of("3"));
    }

    /** The raw {@code showDropDown} attribute: in OOXML {@code true} HIDES the in-cell arrow. */
    private static boolean arrowHidden(XSSFWorkbook wb, int validationIndex) {
        return wb.getSheetAt(0).getCTWorksheet().getDataValidations().getDataValidationArray(validationIndex)
                .getShowDropDown();
    }

    @Nested
    class Arrow {
        @Test
        void everyKindOfListKeepsTheInCellArrow() throws IOException {
            try (var wb = write(ExcelWriter.<Row>create()
                    .column("Fixed", Row::category, c -> c.dropdown("A", "B"))
                    .column("List", Row::category, c -> c.validation(ExcelValidation.list("A", "B")))
                    .column("Range", Row::category, c -> c.validation(ExcelValidation.listFromRange("Z!$A$1:$A$3")))
                    .column("Dependent", Row::item, c -> c.dependentDropdown("Fixed", options())))) {
                assertEquals(4, wb.getSheetAt(0).getDataValidations().size());
                for (int i = 0; i < 4; i++) assertFalse(arrowHidden(wb, i), "validation " + i);
            }
        }
    }

    @Nested
    class ExplicitList {
        @Test
        void takesTheOptionsOfAValidation() throws IOException {
            try (var wb = write(writer(c -> c.validation(ExcelValidation.list("A", "B", "C")
                    .allowBlank(false).inputTitle("Item").inputMessage("Pick one").errorMessage("Not in the list"))))) {
                var v = wb.getSheetAt(0).getDataValidations().get(1);
                assertEquals("\"A,B,C\"", v.getValidationConstraint().getFormula1());
                assertFalse(v.getEmptyCellAllowed());
                assertTrue(v.getShowPromptBox());
                assertEquals("Pick one", v.getPromptBoxText());
                assertEquals("Not in the list", v.getErrorBoxText());
            }
        }

        @Test
        void followsTheRulesOfDropdown() {
            assertThrows(IllegalArgumentException.class, () -> ExcelValidation.list());
            assertThrows(IllegalArgumentException.class, () -> ExcelValidation.list("a,b"));
            assertThrows(IllegalArgumentException.class, () -> ExcelValidation.list("x".repeat(256)));
            assertThrows(NullPointerException.class, () -> ExcelValidation.list((String[]) null));
            assertThrows(NullPointerException.class, () -> ExcelValidation.list("a", null));
        }

        @Test
        void aQuoteIsWrittenTwiceAsForDropdown() throws IOException {
            try (var wb = write(writer(c -> c.validation(ExcelValidation.list("say \"hi\""))))) {
                assertEquals("\"say \"\"hi\"\"\"",
                        wb.getSheetAt(0).getDataValidations().get(1).getValidationConstraint().getFormula1());
            }
        }

        @Test
        void aDropdownAndAListValidationOnOneColumnStillConflict() {
            assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<Row>create()
                    .column("C", Row::category, c -> c.dropdown("A").validation(ExcelValidation.list("B"))));
        }
    }

    @Nested
    class DependentList {
        @Test
        void takesTheOptionsOfAValidation() throws IOException {
            try (var wb = write(writer(c -> c.validation(ExcelValidation.dependentList("Category", options())
                    .allowBlank(false).inputMessage("Pick the item of the category"))))) {
                assertEquals(2, wb.getNumberOfSheets(), "the hidden helper sheet");
                var v = wb.getSheetAt(0).getDataValidations().get(1);
                assertFalse(v.getEmptyCellAllowed());
                assertTrue(v.getShowPromptBox());
                assertEquals("Pick the item of the category", v.getPromptBoxText());
                assertTrue(v.getValidationConstraint().getFormula1().contains("$A2"));
            }
        }

        @Test
        void anUnknownParentColumnFailsOnWrite() {
            assertThrows(ExcelWriteException.class, () -> write(writer(c -> c
                    .validation(ExcelValidation.dependentList("Nope", options())))));
        }

        @Test
        void rejectsInvalidArguments() {
            assertThrows(NullPointerException.class, () -> ExcelValidation.dependentList("P", null));
            assertThrows(IllegalArgumentException.class, () -> ExcelValidation.dependentList(" ", options()));
            assertThrows(IllegalArgumentException.class, () -> ExcelValidation.dependentList("P", Map.of()));
        }
    }
}
