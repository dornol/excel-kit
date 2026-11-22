package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Which text a summary row carries in its label cell. This pins what the code does today: a label given with
 * {@code label(text)} is used when the summary has a single operation; with several operations every row is
 * named after its operation (Sum, Average, ...) and the given text is not written. (The javadoc of
 * {@code label} speaks of "the summary row(s)" without saying so.)
 */
class SummaryLabelRowsTest {

    private record Item(String name, int qty) {
    }

    private static List<String> labels(Consumer<ExcelSummary> summary) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Item>create().column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .summary(summary).write(List.of(new Item("a", 1), new Item("b", 2))).writeTo(out);
        List<String> labels = new ArrayList<>();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFSheet sheet = wb.getSheetAt(0);
            for (int r = 3; r <= sheet.getLastRowNum(); r++) labels.add(sheet.getRow(r).getCell(0).getStringCellValue());
        }
        return labels;
    }

    @Test
    void aSingleOperationUsesTheGivenLabel() throws Exception {
        assertEquals(List.of("Total"), labels(s -> s.label("Total").sum("Qty")));
    }

    @Test
    void aSingleOperationOverSeveralColumnsStillUsesTheGivenLabel() throws Exception {
        assertEquals(List.of("Total"), labels(s -> s.label("Total").sum("Qty").sum("Qty")));
    }

    @Test
    void severalOperationsAreNamedAfterTheOperationAndNotAfterTheGivenLabel() throws Exception {
        assertEquals(List.of("Sum", "Average", "Max"), labels(s -> s.label("Total").sum("Qty").average("Qty").max("Qty")));
    }

    @Test
    void withoutALabelASingleOperationIsNamedAfterItsOperation() throws Exception {
        assertEquals(List.of("Sum"), labels(s -> s.sum("Qty")));
    }
}
