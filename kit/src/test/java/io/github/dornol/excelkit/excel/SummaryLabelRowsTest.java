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
 * Which text a summary row carries in its label cell. A label given with {@code label(text)} is used as it is when the
 * summary has a single operation; with several operations every row needs its own name, so the label is followed by
 * the operation ("Total - Sum"). Without a label a row is named after its operation.
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
    void severalOperationsKeepTheGivenLabelAndAddTheOperation() throws Exception {
        assertEquals(List.of("Total - Sum", "Total - Average", "Total - Max"),
                labels(s -> s.label("Total").sum("Qty").average("Qty").max("Qty")));
    }

    @Test
    void theLabelOfAColumnIsExtendedTheSameWay() throws Exception {
        assertEquals(List.of("Grand - Sum", "Grand - Min"), labels(s -> s.label("Name", "Grand").sum("Qty").min("Qty")));
        assertEquals(List.of("Grand"), labels(s -> s.label("Name", "Grand").sum("Qty")));
    }

    @Test
    void withoutALabelSeveralOperationsAreNamedAfterTheOperation() throws Exception {
        assertEquals(List.of("Sum", "Average"), labels(s -> s.sum("Qty").average("Qty")));
    }

    @Test
    void withoutALabelASingleOperationIsNamedAfterItsOperation() throws Exception {
        assertEquals(List.of("Sum"), labels(s -> s.sum("Qty")));
    }
}
