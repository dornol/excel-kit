package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
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

/** The same column name is allowed in different header groups, and rejected where the name is the only key. */
class SameNameInDifferentGroupsTest {

    private record Row(String a, String b) {}

    private static XSSFWorkbook open(ExcelHandler handler) throws IOException {
        var out = new ByteArrayOutputStream();
        handler.writeTo(out);
        return new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
    }

    private static String text(XSSFSheet sheet, int row, int col) {
        var cell = sheet.getRow(row).getCell(col);
        return cell == null ? "" : cell.getStringCellValue();
    }

    private static Stream<Row> one() { return Stream.of(new Row("1", "2")); }

    @Nested
    class DuplicateCheck {
        @Test
        void differentGroupsWriteBothColumnsInPlace() throws IOException {
            try (var wb = open(ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("Home"))
                    .column("Name", Row::b, c -> c.group("Work"))
                    .write(one()))) {
                var sheet = wb.getSheetAt(0);
                assertEquals("Home", text(sheet, 0, 0));
                assertEquals("Work", text(sheet, 0, 1));
                assertEquals("Name", text(sheet, 1, 0));
                assertEquals("Name", text(sheet, 1, 1));
                assertEquals("1", text(sheet, 2, 0));
                assertEquals("2", text(sheet, 2, 1));
            }
        }

        @Test
        void groupedAndUngroupedWithTheSameNameAreDifferent() {
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("Name", Row::a)
                    .column("Name", Row::b, c -> c.group("Work"))
                    .write(one())));
        }

        @Test
        void sameGroupStillFails() {
            var writer = ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("Home"))
                    .column("Name", Row::b, c -> c.group("Home"));
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("Duplicate column name: 'Name'"), error.getMessage());
            assertTrue(error.getMessage().contains("Home"), "names the group: " + error.getMessage());
        }

        @Test
        void ungroupedDuplicateStillFails() {
            var writer = ExcelWriter.<Row>create().column("Name", Row::a).column("Name", Row::b);
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertFalse(error.getMessage().contains(" in group"), error.getMessage());
        }

        @Test
        void aDifferenceInOnlyTheOuterLevelIsEnough() {
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("X", "Inner"))
                    .column("Name", Row::b, c -> c.group("Y", "Inner"))
                    .write(one())));
        }

        @Test
        void aDifferenceInOnlyTheInnerLevelIsEnough() {
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("Outer", "X"))
                    .column("Name", Row::b, c -> c.group("Outer", "Y"))
                    .write(one())));
        }

        @Test
        void aGroupPathThatIsAPrefixIsADifferentGroup() {
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("Outer"))
                    .column("Name", Row::b, c -> c.group("Outer", "Inner"))
                    .write(one())));
        }

        @Test
        void samePathOfSeveralLevelsFails() {
            var writer = ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("Outer", "Inner"))
                    .column("Name", Row::b, c -> c.group("Outer", "Inner"));
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("Outer > Inner"), error.getMessage());
        }

        @Test
        void aGroupNameIsNotMistakenForAColumnName() {
            // Column "B" under group "A" and column "A > B" are different keys; a path is not flattened into a string.
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("B", Row::a, c -> c.group("A"))
                    .column("A", Row::b, c -> c.group("B"))
                    .write(one())));
        }

        @Test
        void aNonAdjacentSameGroupStillFails() {
            var writer = ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("Home"))
                    .column("Other", Row::b, c -> c.group("Work"))
                    .column("Name", Row::a, c -> c.group("Home"));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void workbookSheetWritersFollowTheSameRule() {
            try (ExcelWorkbook wb = ExcelWorkbook.create()) {
                var ok = wb.<Row>sheet("Ok")
                        .column("Name", Row::a, c -> c.group("Home"))
                        .column("Name", Row::b, c -> c.group("Work"));
                assertDoesNotThrow(() -> ok.write(one()));
                var bad = wb.<Row>sheet("Bad")
                        .column("Name", Row::a, c -> c.group("Home"))
                        .column("Name", Row::b, c -> c.group("Home"));
                assertThrows(ExcelWriteException.class, () -> bad.write(Stream.of(new Row("1", "2"))));
            }
        }
    }

    @Nested
    class InATable {
        @Test
        void aTableRejectsTheSameHeaderEvenInDifferentGroups() {
            var writer = ExcelWriter.<Row>create().table("People")
                    .column("Name", Row::a, c -> c.group("Home"))
                    .column("Name", Row::b, c -> c.group("Work"));
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("Duplicate table header"), error.getMessage());
        }

        @Test
        void aTableComparesHeadersIgnoringCase() {
            var writer = ExcelWriter.<Row>create().table("People")
                    .column("Name", Row::a, c -> c.group("Home"))
                    .column("NAME", Row::b, c -> c.group("Work"));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void aTableWithDistinctHeadersUnderGroupsStillWorks() throws IOException {
            try (var wb = open(ExcelWriter.<Row>create().table("People")
                    .column("First", Row::a, c -> c.group("Home"))
                    .column("Second", Row::b, c -> c.group("Work"))
                    .write(one()))) {
                assertEquals(1, wb.getSheetAt(0).getTables().size());
            }
        }

        @Test
        void theSameNamesWithoutATableAreFine() {
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("Home"))
                    .column("Name", Row::b, c -> c.group("Work"))
                    .write(one())));
        }
    }

    @Nested
    class Subtotal {
        private ExcelWriter<Row> writer(String sumColumn) {
            return ExcelWriter.<Row>create()
                    .column("Key", Row::a)
                    .column("Amount", Row::b, c -> c.group("Q1"))
                    .column("Amount", Row::b, c -> c.group("Q2"))
                    .groupSubtotal(Row::a, s -> s.label("Key", k -> "sub").sum(sumColumn));
        }

        @Test
        void aNameSharedByTwoColumnsIsAmbiguous() {
            var writer = writer("Amount");
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("more than one column"), error.getMessage());
        }

        @Test
        void aUniqueNameNextToDuplicatesStillResolves() {
            var writer = ExcelWriter.<Row>create()
                    .column("Key", Row::a)
                    .column("Amount", Row::b, c -> c.group("Q1"))
                    .column("Amount", Row::b, c -> c.group("Q2"))
                    .column("Total", Row::b)
                    .groupSubtotal(Row::a, s -> s.label("Key", k -> "sub").sum("Total"));
            assertDoesNotThrow(() -> open(writer.write(one())));
        }

        @Test
        void anAmbiguousLabelColumnIsRejected() {
            var writer = ExcelWriter.<Row>create()
                    .column("Label", Row::a, c -> c.group("X"))
                    .column("Label", Row::a, c -> c.group("Y"))
                    .column("Amount", Row::b)
                    .groupSubtotal(Row::a, s -> s.label("Label", k -> "sub").sum("Amount"));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void anAmbiguousKeyColumnIsRejected() {
            var writer = ExcelWriter.<Row>create()
                    .column("Label", Row::a)
                    .column("Key", Row::a, c -> c.group("X"))
                    .column("Key", Row::a, c -> c.group("Y"))
                    .column("Amount", Row::b)
                    .groupSubtotal(Row::a, s -> s.label("Label", k -> "sub").keyColumn("Key").sum("Amount"));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void aMissingNameStillReportsNotFound() {
            var writer = writer("Nope");
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("not found"), error.getMessage());
        }
    }

    @Nested
    class Summary {
        private ExcelWriter<Row> base() {
            return ExcelWriter.<Row>create()
                    .column("Label", Row::a)
                    .column("Amount", Row::b, c -> c.group("Q1"))
                    .column("Amount", Row::b, c -> c.group("Q2"))
                    .column("Total", Row::b);
        }

        @Test
        void aSummaryOverANameSharedByTwoColumnsIsAmbiguous() {
            var writer = base().summary(s -> s.label("Label", "Sum").sum("Amount"));
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("more than one column"), error.getMessage());
        }

        @Test
        void anAmbiguousSummaryLabelColumnIsRejected() {
            var writer = ExcelWriter.<Row>create()
                    .column("Label", Row::a, c -> c.group("X"))
                    .column("Label", Row::a, c -> c.group("Y"))
                    .column("Total", Row::b)
                    .summary(s -> s.label("Label", "Sum").sum("Total"));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void aUniqueNameNextToDuplicatesStillSummarises() throws IOException {
            try (var wb = open(base().summary(s -> s.label("Label", "Sum").sum("Total")).write(one()))) {
                var sheet = wb.getSheetAt(0);
                var formula = sheet.getRow(sheet.getLastRowNum()).getCell(3).getCellFormula();
                assertTrue(formula.contains("D"), formula);
            }
        }
    }

    @Nested
    class ColumnId {
        private ExcelWriter<Row> twoAmounts() {
            return ExcelWriter.<Row>create()
                    .column("Label", Row::a)
                    .column("Amount", Row::b, c -> c.group("Q1").id("q1"))
                    .column("Amount", Row::b, c -> c.group("Q2").id("q2"));
        }

        private static String lastRowFormula(XSSFWorkbook wb, int col) {
            var sheet = wb.getSheetAt(0);
            return sheet.getRow(sheet.getLastRowNum()).getCell(col).getCellFormula();
        }

        @Test
        void aSummaryReferringToAnIdPicksThatColumn() throws IOException {
            try (var wb = open(twoAmounts().summary(s -> s.label("Label", "Sum").sum("q2")).write(one()))) {
                assertTrue(lastRowFormula(wb, 2).startsWith("SUM(C"), lastRowFormula(wb, 2));
                var last = wb.getSheetAt(0).getRow(wb.getSheetAt(0).getLastRowNum());
                var other = last.getCell(1);
                assertTrue(other == null || other.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK,
                        "the other Amount is not summed");
            }
        }

        @Test
        void bothColumnsCanBeSummedByTheirIds() throws IOException {
            try (var wb = open(twoAmounts().summary(s -> s.label("Label", "Sum").sum("q1").sum("q2")).write(one()))) {
                var sheet = wb.getSheetAt(0);
                var row = sheet.getRow(sheet.getLastRowNum());
                assertTrue(row.getCell(1).getCellFormula().startsWith("SUM(B"));
                assertTrue(row.getCell(2).getCellFormula().startsWith("SUM(C"));
            }
        }

        @Test
        void aSubtotalReferringToAnIdPicksThatColumn() {
            var writer = twoAmounts().groupSubtotal(Row::a, s -> s.label("Label", k -> "sub").sum("q1"));
            assertDoesNotThrow(() -> open(writer.write(one())));
        }

        @Test
        void aNameStillWorksWhenItIsUnique() {
            var writer = ExcelWriter.<Row>create()
                    .column("Label", Row::a)
                    .column("Total", Row::b, c -> c.id("t"))
                    .summary(s -> s.label("Label", "Sum").sum("Total"));
            assertDoesNotThrow(() -> open(writer.write(one())));
        }

        @Test
        void anAmbiguousNameTellsToUseIds() {
            var writer = twoAmounts().summary(s -> s.label("Label", "Sum").sum("Amount"));
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("id"), error.getMessage());
        }

        @Test
        void theSameNameInTheSameGroupIsAllowedWhenBothHaveIds() {
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("Amount", Row::a, c -> c.group("Q").id("x"))
                    .column("Amount", Row::b, c -> c.group("Q").id("y"))
                    .write(one())));
        }

        @Test
        void theSameNameInTheSameGroupNeedsAnIdOnBoth() {
            var onlyOne = ExcelWriter.<Row>create()
                    .column("Amount", Row::a, c -> c.group("Q").id("x"))
                    .column("Amount", Row::b, c -> c.group("Q"));
            assertThrows(ExcelWriteException.class, () -> onlyOne.write(one()));
        }

        @Test
        void anIdMustBeUnique() {
            var writer = ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id("k"))
                    .column("B", Row::b, c -> c.id("k"));
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("Duplicate column id"), error.getMessage());
        }

        @Test
        void anIdMayNotBeTheNameOfAnotherColumn() {
            var writer = ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id("B"))
                    .column("B", Row::b);
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void anIdMayBeTheNameOfItsOwnColumn() {
            assertDoesNotThrow(() -> open(ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id("A"))
                    .write(one())));
        }

        @Test
        void aMissingReferenceStillReportsNotFound() {
            var writer = twoAmounts().summary(s -> s.label("Label", "Sum").sum("nope"));
            var error = assertThrows(ExcelWriteException.class, () -> writer.write(one()));
            assertTrue(error.getMessage().contains("not found"), error.getMessage());
        }

        private static ExcelWriter<Row> numbers() {
            // column B holds 10 and column C holds 5 so that a summary of the wrong column shows
            return ExcelWriter.<Row>create()
                    .column("Label", Row::a)
                    .column("Amount", r -> 10, c -> c.group("Q1").id("q1"))
                    .column("Amount", r -> 5, c -> c.group("Q2").id("q2"));
        }

        private static double cachedValue(XSSFWorkbook wb, int col) {
            var sheet = wb.getSheetAt(0);
            var cell = sheet.getRow(sheet.getLastRowNum()).getCell(col);
            return new org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator(wb).evaluate(cell).getNumberValue();
        }

        @Test
        void aSummaryOfAnIdCalculatesThatColumn() throws IOException {
            try (var wb = open(numbers().summary(s -> s.label("Label", "Sum").sum("q2"))
                    .write(Stream.of(new Row("x", "y"), new Row("x", "y"))))) {
                assertEquals(10, cachedValue(wb, 2), "two rows of 5");
            }
        }

        @Test
        void anAverageOfAnIdCalculatesThatColumn() throws IOException {
            try (var wb = open(numbers().summary(s -> s.label("Label", "Avg").average("q1"))
                    .write(Stream.of(new Row("x", "y"), new Row("x", "y"))))) {
                assertEquals(10, cachedValue(wb, 1));
            }
        }

        @Test
        void aSummaryLabelColumnCanBeGivenById() throws IOException {
            var writer = ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("X").id("first"))
                    .column("Name", Row::b, c -> c.group("Y").id("second"))
                    .column("Total", r -> 3)
                    .summary(s -> s.label("second", "Sum").sum("Total"));
            try (var wb = open(writer.write(one()))) {
                var sheet = wb.getSheetAt(0);
                var last = sheet.getRow(sheet.getLastRowNum());
                assertEquals("Sum", last.getCell(1).getStringCellValue(), "the label sits in the column of the id");
            }
        }

        @Test
        void aSubtotalOfAnIdCalculatesThatColumn() throws IOException {
            var writer = numbers().groupSubtotal(Row::a, s -> s.label("Label", k -> "sub").sum("q2"));
            try (var wb = open(writer.write(Stream.of(new Row("x", "y"), new Row("x", "y"))))) {
                var sheet = wb.getSheetAt(0);
                var subtotal = sheet.getRow(sheet.getLastRowNum());
                var eval = new org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator(wb);
                assertEquals(10, eval.evaluate(subtotal.getCell(2)).getNumberValue());
                assertTrue(subtotal.getCell(1) == null
                        || subtotal.getCell(1).getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK);
            }
        }

        @Test
        void aSubtotalLabelAndKeyColumnCanBeGivenById() {
            var writer = ExcelWriter.<Row>create()
                    .column("Name", Row::a, c -> c.group("X").id("label"))
                    .column("Name", Row::a, c -> c.group("Y").id("key"))
                    .column("Total", r -> 3)
                    .groupSubtotal(Row::a, s -> s.label("label", k -> "sub").keyColumn("key").sum("Total"));
            assertDoesNotThrow(() -> open(writer.write(one())));
        }

        @Test
        void aSubtotalRejectsTheLabelColumnAsASumColumnById() {
            var subtotal = numbers().groupSubtotal(Row::a, s -> s.label("q1", k -> "sub").sum("q1"));
            assertThrows(ExcelWriteException.class, () -> subtotal.write(one()));
        }

        @Test
        void sheetWritersOfAWorkbookTakeIdsToo() throws IOException {
            try (ExcelWorkbook wb = ExcelWorkbook.create()) {
                var sheet = wb.<Row>sheet("S")
                        .column("Label", Row::a)
                        .column("Amount", r -> 10, c -> c.group("Q1").id("q1"))
                        .column("Amount", r -> 5, c -> c.group("Q1").id("q2"))
                        .summary(s -> s.label("Label", "Sum").sum("q2"));
                assertDoesNotThrow(() -> sheet.write(one()));
            }
        }

        @Test
        void aTemplateListTakesIdsToo() throws Exception {
            byte[] template;
            try (var book = new XSSFWorkbook()) {
                book.createSheet("S").createRow(0).createCell(0).setCellValue("x");
                var out = new ByteArrayOutputStream();
                book.write(out);
                template = out.toByteArray();
            }
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
                var list = tpl.<Row>list(1)
                        .column("Amount", Row::a, c -> c.group("Q").id("x"))
                        .column("Amount", Row::b, c -> c.group("Q").id("y"));
                assertDoesNotThrow(() -> list.write(one()));
            }
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
                var list = tpl.<Row>list(1)
                        .column("Amount", Row::a, c -> c.group("Q").id("x"))
                        .column("Amount", Row::b, c -> c.group("Q").id("x"));
                assertThrows(ExcelWriteException.class, () -> list.write(one()));
            }
        }

        @Test
        void threeColumnsOfTheSameNameNeedAnIdOnEveryOne() {
            var allIds = ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id("1"))
                    .column("A", Row::a, c -> c.id("2"))
                    .column("A", Row::a, c -> c.id("3"));
            assertDoesNotThrow(() -> open(allIds.write(one())));
            var oneMissing = ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id("1"))
                    .column("A", Row::a, c -> c.id("2"))
                    .column("A", Row::a);
            assertThrows(ExcelWriteException.class, () -> oneMissing.write(one()));
            var firstMissing = ExcelWriter.<Row>create()
                    .column("A", Row::a)
                    .column("A", Row::a, c -> c.id("2"))
                    .column("A", Row::a, c -> c.id("3"));
            assertThrows(ExcelWriteException.class, () -> firstMissing.write(one()));
        }

        @Test
        void anIdOfALaterColumnMayNotBeTheNameOfAnEarlierOne() {
            var writer = ExcelWriter.<Row>create()
                    .column("B", Row::b)
                    .column("A", Row::a, c -> c.id("B"));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void anIdIsCaseSensitive() {
            var writer = ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id("k"))
                    .column("B", Row::b, c -> c.id("K"));
            assertDoesNotThrow(() -> open(writer.write(one())));
        }

        @Test
        void aColumnIsNeverItsOwnParentById() {
            var writer = ExcelWriter.<Row>create()
                    .column("Item", Row::b, c -> c.id("me").dependentDropdown("me", Map.of("A", List.of("1"))));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void aDefaultStyleCannotCarryAnId() {
            var writer = ExcelWriter.<Row>create();
            assertThrows(UnsupportedOperationException.class, () -> writer.defaultStyle(d -> d.id("x")));
        }

        @Test
        void columnsThatDoNotSetAnIdAreNotGivenOneByTheDefaults() {
            var writer = ExcelWriter.<Row>create()
                    .defaultStyle(d -> d.bold(true))
                    .column("A", Row::a, c -> c.id("a"))
                    .column("B", Row::b)
                    .summary(s -> s.label("A", "Sum").sum("B"));
            assertDoesNotThrow(() -> open(writer.write(one())));
        }

        @Test
        void aDependentDropdownFindsItsParentById() throws IOException {
            var writer = ExcelWriter.<Row>create()
                    .column("Cat", Row::a, c -> c.group("X").id("catX").dropdown("A"))
                    .column("Cat", Row::a, c -> c.group("Y").id("catY").dropdown("A"))
                    .column("Item", Row::b, c -> c.group("Y").dependentDropdown("catX", Map.of("A", List.of("1"))));
            try (var wb = open(writer.write(one()))) {
                String f = wb.getSheetAt(0).getDataValidations().get(2).getValidationConstraint().getFormula1();
                assertTrue(f.contains("$A3"), "the parent is column A, not the same-group Cat: " + f);
            }
        }

        @Test
        void aTableStillRejectsTheSameHeaderEvenWithIds() {
            var writer = ExcelWriter.<Row>create().table("People")
                    .column("Name", Row::a, c -> c.group("Home").id("h"))
                    .column("Name", Row::b, c -> c.group("Work").id("w"));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void aBlankOrNullIdIsRejected() {
            assertThrows(IllegalArgumentException.class, () -> ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id(" ")));
            assertThrows(NullPointerException.class, () -> ExcelWriter.<Row>create()
                    .column("A", Row::a, c -> c.id(null)));
        }
    }

    @Nested
    class Rollover {
        private record Num(String key, int x, int y) {}

        /** Five rows with maxRows(2): three sheets of 2, 2 and 1 data rows. */
        private static Stream<Num> five() {
            return Stream.of(new Num("k", 1, 10), new Num("k", 2, 20), new Num("k", 3, 30),
                    new Num("k", 4, 40), new Num("k", 5, 50));
        }

        private static ExcelWriter<Num> twoAmounts() {
            return ExcelWriter.<Num>create().maxRows(2)
                    .column("Key", Num::key)
                    .column("Amount", Num::x, c -> c.group("Q1").id("q1"))
                    .column("Amount", Num::y, c -> c.group("Q2").id("q2"));
        }

        private static double value(XSSFWorkbook wb, XSSFSheet sheet, int row, int col) {
            return new org.apache.poi.xssf.usermodel.XSSFFormulaEvaluator(wb)
                    .evaluate(sheet.getRow(row).getCell(col)).getNumberValue();
        }

        @Test
        void everySheetRepeatsTheGroupedHeadersOfTheSameName() throws IOException {
            try (var wb = open(twoAmounts().write(five()))) {
                assertEquals(3, wb.getNumberOfSheets());
                for (int i = 0; i < 3; i++) {
                    var sheet = wb.getSheetAt(i);
                    assertEquals("Q1", text(sheet, 0, 1), "sheet " + i);
                    assertEquals("Q2", text(sheet, 0, 2), "sheet " + i);
                    assertEquals("Amount", text(sheet, 1, 1), "sheet " + i);
                    assertEquals("Amount", text(sheet, 1, 2), "sheet " + i);
                }
            }
        }

        @Test
        void aSummaryOfAnIdCalculatesTheRightColumnOnEverySheet() throws IOException {
            var writer = twoAmounts().summary(s -> s.label("Key", "Sum").sum("q2"));
            try (var wb = open(writer.write(five()))) {
                // data rows 1+2 | 3+4 | 5 on the Q2 column (10,20 | 30,40 | 50)
                double[] expected = {30, 70, 50};
                for (int i = 0; i < 3; i++) {
                    var sheet = wb.getSheetAt(i);
                    int last = sheet.getLastRowNum();
                    assertEquals(expected[i], value(wb, sheet, last, 2), "sheet " + i);
                    var other = sheet.getRow(last).getCell(1);
                    assertTrue(other == null || other.getCellType() == org.apache.poi.ss.usermodel.CellType.BLANK,
                            "the Q1 Amount is not summed on sheet " + i);
                }
            }
        }

        @Test
        void aSubtotalOfAnIdCalculatesTheRightColumnOnEverySheet() throws IOException {
            var writer = twoAmounts().groupSubtotal(Num::key, s -> s.label("Key", k -> "sub").sum("q1"));
            try (var wb = open(writer.write(five()))) {
                assertEquals(3, wb.getNumberOfSheets());
                double[] expected = {3, 7, 5};
                for (int i = 0; i < 3; i++) {
                    var sheet = wb.getSheetAt(i);
                    int last = sheet.getLastRowNum();
                    assertEquals(expected[i], value(wb, sheet, last, 1), "sheet " + i);
                }
            }
        }

        @Test
        void aDependentDropdownOfAnIdPointsAtTheRightColumnOnEverySheet() throws IOException {
            var writer = ExcelWriter.<Num>create().maxRows(2)
                    .column("Cat", Num::key, c -> c.group("X").id("catX").dropdown("k"))
                    .column("Cat", Num::key, c -> c.group("Y").id("catY").dropdown("k"))
                    .column("Item", Num::key, c -> c.group("Y").dependentDropdown("catX", Map.of("k", List.of("1"))));
            try (var wb = open(writer.write(five()))) {
                for (int i = 0; i < 3; i++) {
                    var validations = wb.getSheetAt(i).getDataValidations();
                    String f = validations.get(validations.size() - 1).getValidationConstraint().getFormula1();
                    assertTrue(f.contains("$A3"), "sheet " + i + " refers to column A (catX): " + f);
                }
            }
        }

        @Test
        void aDuplicateInTheSameGroupFailsBeforeAnySheetIsWritten() {
            var writer = ExcelWriter.<Num>create().maxRows(2)
                    .column("Amount", Num::x, c -> c.group("Q"))
                    .column("Amount", Num::y, c -> c.group("Q"));
            assertThrows(ExcelWriteException.class, () -> writer.write(five()));
        }

        @Test
        void aTableOverEverySheetStillRejectsTheSameHeader() {
            var writer = ExcelWriter.<Num>create().maxRows(2).table("People")
                    .column("Amount", Num::x, c -> c.group("Q1").id("q1"))
                    .column("Amount", Num::y, c -> c.group("Q2").id("q2"));
            assertThrows(ExcelWriteException.class, () -> writer.write(five()));
        }

        @Test
        void aTableOfDistinctHeadersGetsOneTablePerSheet() throws IOException {
            var writer = ExcelWriter.<Num>create().maxRows(2).table("People")
                    .column("Key", Num::key)
                    .column("First", Num::x, c -> c.group("Q1").id("q1"))
                    .column("Second", Num::y, c -> c.group("Q2").id("q2"));
            try (var wb = open(writer.write(five()))) {
                for (int i = 0; i < 3; i++) assertEquals(1, wb.getSheetAt(i).getTables().size(), "sheet " + i);
            }
        }

        @Test
        void aWorkbookSheetWriterRollsOverWithIdsToo() throws IOException {
            try (ExcelWorkbook book = ExcelWorkbook.create()) {
                var sheet = book.<Num>sheet("Data").maxRows(2)
                        .column("Key", Num::key)
                        .column("Amount", Num::x, c -> c.group("Q1").id("q1"))
                        .column("Amount", Num::y, c -> c.group("Q2").id("q2"))
                        .summary(s -> s.label("Key", "Sum").sum("q2"));
                sheet.write(five());
                try (var wb = open(book.finish())) {
                    assertEquals(3, wb.getNumberOfSheets());
                    assertEquals(50, value(wb, wb.getSheetAt(2), wb.getSheetAt(2).getLastRowNum(), 2));
                }
            }
        }
    }

    @Nested
    class DependentDropdown {
        private static Map<String, List<String>> options() {
            return Map.of("A", List.of("1", "2"));
        }

        private static String formula(ExcelWriter<Row> writer, int validation) throws IOException {
            try (var wb = open(writer.write(one()))) {
                return wb.getSheetAt(0).getDataValidations().get(validation).getValidationConstraint().getFormula1();
            }
        }

        @Test
        void theParentInTheSameGroupWins() throws IOException {
            var writer = ExcelWriter.<Row>create()
                    .column("Cat", Row::a, c -> c.group("Other").dropdown("A"))
                    .column("Cat", Row::a, c -> c.group("Mine").dropdown("A"))
                    .column("Item", Row::b, c -> c.group("Mine").dependentDropdown("Cat", options()));
            // the parent is column B (index 1), the one of group "Mine"
            String f = formula(writer, 2);
            assertTrue(f.contains("$B3"), f);
        }

        @Test
        void theParentOfTheSameGroupWinsEvenWhenItComesLater() throws IOException {
            var writer = ExcelWriter.<Row>create()
                    .column("Cat", Row::a, c -> c.group("Other").dropdown("A"))
                    .column("Item", Row::b, c -> c.group("Mine").dependentDropdown("Cat", options()))
                    .column("Cat", Row::a, c -> c.group("Mine").dropdown("A"));
            String f = formula(writer, 1);
            assertTrue(f.contains("$C3"), f);
        }

        @Test
        void aParentOfAnotherGroupIsUsedWhenNoneIsInTheSameGroup() throws IOException {
            var writer = ExcelWriter.<Row>create()
                    .column("Cat", Row::a, c -> c.group("Other").dropdown("A"))
                    .column("Item", Row::b, c -> c.group("Mine").dependentDropdown("Cat", options()));
            String f = formula(writer, 1);
            assertTrue(f.contains("$A3"), f);
        }

        @Test
        void aColumnIsNeverItsOwnParent() {
            var writer = ExcelWriter.<Row>create()
                    .column("Item", Row::b, c -> c.dependentDropdown("Item", options()));
            assertThrows(ExcelWriteException.class, () -> writer.write(one()));
        }

        @Test
        void aSameNamedColumnOfAnotherGroupIsAValidParent() throws IOException {
            // "Item" depends on the "Item" of the other group, never on itself
            var writer = ExcelWriter.<Row>create()
                    .column("Item", Row::a, c -> c.group("X").dropdown("A"))
                    .column("Item", Row::b, c -> c.group("Y").dependentDropdown("Item", options()));
            String f = formula(writer, 1);
            assertTrue(f.contains("$A3"), f);
        }
    }

    @Nested
    class TemplateList {
        private static byte[] template() throws IOException {
            try (var wb = new XSSFWorkbook()) {
                var sheet = wb.createSheet("S");
                sheet.createRow(0).createCell(0).setCellValue("x");
                var out = new ByteArrayOutputStream();
                wb.write(out);
                return out.toByteArray();
            }
        }

        @Test
        void differentGroupsAreAllowedAndTheSameGroupFails() throws Exception {
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var ok = tpl.<Row>list(2)
                        .column("Name", Row::a, c -> c.group("Home"))
                        .column("Name", Row::b, c -> c.group("Work"));
                assertDoesNotThrow(() -> ok.write(Stream.of(new Row("1", "2"))));
            }
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var bad = tpl.<Row>list(2)
                        .column("Name", Row::a, c -> c.group("Home"))
                        .column("Name", Row::b, c -> c.group("Home"));
                var error = assertThrows(ExcelWriteException.class, () -> bad.write(Stream.of(new Row("1", "2"))));
                assertTrue(error.getMessage().contains("Duplicate column name"), error.getMessage());
            }
        }

        /** A template whose header row (row 1, 0-based) holds the given headers. */
        private static byte[] templateWithHeaders(String... headers) throws IOException {
            try (var wb = new XSSFWorkbook()) {
                var sheet = wb.createSheet("S");
                var row = sheet.createRow(1);
                for (int i = 0; i < headers.length; i++) row.createCell(i).setCellValue(headers[i]);
                var out = new ByteArrayOutputStream();
                wb.write(out);
                return out.toByteArray();
            }
        }

        @Test
        void aTableOverDuplicateHeadersOfTheTemplateIsRejected() throws Exception {
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(templateWithHeaders("Name", "Name")))) {
                var list = tpl.<Row>list(2)
                        .column("a", Row::a).column("b", Row::b).table("Items");
                var error = assertThrows(ExcelWriteException.class, () -> list.write(Stream.of(new Row("1", "2"))));
                assertTrue(error.getMessage().contains("Duplicate table header"), error.getMessage());
            }
        }

        @Test
        void theTemplateHeadersAreComparedIgnoringCase() throws Exception {
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(templateWithHeaders("Name", "NAME")))) {
                var list = tpl.<Row>list(2)
                        .column("a", Row::a).column("b", Row::b).table("Items");
                assertThrows(ExcelWriteException.class, () -> list.write(Stream.of(new Row("1", "2"))));
            }
        }

        @Test
        void distinctTemplateHeadersMakeATable() throws Exception {
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(templateWithHeaders("Name", "Age")))) {
                tpl.<Row>list(2).column("a", Row::a).column("b", Row::b).table("Items")
                        .write(Stream.of(new Row("1", "2")));
                try (var wb = open(tpl.finish())) {
                    assertEquals(1, wb.getSheetAt(0).getTables().size());
                }
            }
        }

        @Test
        void duplicateTemplateHeadersWithoutATableAreFine() throws Exception {
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(templateWithHeaders("Name", "Name")))) {
                assertDoesNotThrow(() -> tpl.<Row>list(2).column("a", Row::a).column("b", Row::b)
                        .write(Stream.of(new Row("1", "2"))));
            }
        }

        @Test
        void aListThatWritesItsOwnHeadersRejectsTheSameHeaderInATable() throws Exception {
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var list = tpl.<Row>list(2).table("Items")
                        .column("Name", Row::a, c -> c.group("Home"))
                        .column("Name", Row::b, c -> c.group("Work"));
                var error = assertThrows(ExcelWriteException.class,
                        () -> list.writeWithHeaders(Stream.of(new Row("1", "2"))));
                assertTrue(error.getMessage().contains("Duplicate table header"), error.getMessage());
            }
        }

        @Test
        void aListThatWritesItsOwnHeadersAllowsTheSameNameOutsideATable() throws Exception {
            try (var tpl = new ExcelTemplateWriter(new ByteArrayInputStream(template()))) {
                var list = tpl.<Row>list(2)
                        .column("Name", Row::a, c -> c.group("Home"))
                        .column("Name", Row::b, c -> c.group("Work"));
                assertDoesNotThrow(() -> list.writeWithHeaders(Stream.of(new Row("1", "2"))));
            }
        }
    }
}
