package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Set;
import java.util.TreeSet;

import static org.junit.jupiter.api.Assertions.assertEquals;

class GroupHeaderMergeTest {

    private static Set<String> mergedRegions(ByteArrayOutputStream out) throws IOException {
        try (var workbook = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var sheet = workbook.getSheetAt(0);
            Set<String> regions = new TreeSet<>();
            for (int i = 0; i < sheet.getNumMergedRegions(); i++) {
                regions.add(sheet.getMergedRegion(i).formatAsString());
            }
            return regions;
        }
    }

    @Test
    void equalInnerLabelsUnderDifferentParentsAreNotMergedTogether() throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Net", s -> s, c -> c.group("Sales", "Total"))
                .column("Tax", s -> s, c -> c.group("Sales", "Total"))
                .column("Fee", s -> s, c -> c.group("Cost", "Total"))
                .write(List.of("a")).writeTo(out);

        // Sales spans A:B; the "Total" under it spans only A:B, and Cost/Total stays a single cell.
        assertEquals(Set.of("A1:B1", "A2:B2"), mergedRegions(out));
    }

    @Test
    void equalLabelsUnderTheSameParentStillMerge() throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s, c -> c.group("Top", "Mid", "Leaf"))
                .column("B", s -> s, c -> c.group("Top", "Mid", "Leaf"))
                .column("C", s -> s, c -> c.group("Top", "Other", "Leaf"))
                .write(List.of("a")).writeTo(out);

        // Top spans all three; Mid and Leaf merge only for A:B, which share the same parents.
        assertEquals(Set.of("A1:C1", "A2:B2", "A3:B3"), mergedRegions(out));
    }
}
