package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.RowData;
import org.apache.poi.util.LocaleUtil;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Locale;
import java.util.Objects;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * {@code fillMergedCells()} on a workbook that Microsoft Excel itself saved after the regions were merged by hand (see
 * {@code external-workbooks/README.md}). The other merged-cell tests use files written by POI.
 */
class ExcelSavedMergedCellsTest {

    @BeforeEach
    void fixedFormattingLocale() {
        LocaleUtil.setUserLocale(Locale.US);
    }

    @AfterEach
    void resetLocale() {
        LocaleUtil.resetUserLocale();
    }

    private static Path fixture() throws Exception {
        return Path.of(Objects.requireNonNull(ExcelSavedMergedCellsTest.class
                .getResource("/external-workbooks/excel-merged-cells.xlsx")).toURI());
    }

    private static List<RowData> read(int sheet, boolean fill) throws Exception {
        var rows = new ArrayList<RowData>();
        var reader = ExcelReader.<RowData>mapping(row -> row).sheetIndex(sheet)
                .duplicateHeaderPolicy(io.github.dornol.excelkit.core.read.DuplicateHeaderPolicy.FIRST);
        if (fill) reader.fillMergedCells();
        reader.read(fixture(), r -> { assertTrue(r.success(), r.messages().toString()); rows.add(r.data()); });
        return rows;
    }

    private static List<String> column(List<RowData> rows, String header) {
        return rows.stream().map(r -> r.get(header).asString()).toList();
    }

    @Test
    void excelLeavesTheCoveredCellsEmptyAndWithoutTheOptionTheyReadBlank() throws Exception {
        var rows = read(0, false);
        assertEquals(6, rows.size());
        assertEquals(List.of("Sales", "", "", "Dev", "", "Ops"), column(rows, "Dept"));
        assertEquals(List.of("Ann", "Bob", "Cy", "Di", "Ed", "Flo"), column(rows, "Name"));
    }

    @Test
    void everyRowOfARegionMergedByExcelReadsTheValue() throws Exception {
        var rows = read(0, true);
        assertEquals(List.of("Sales", "Sales", "Sales", "Dev", "Dev", "Ops"), column(rows, "Dept"));
        assertEquals(List.of("Ann", "Bob", "Cy", "Di", "Ed", "Flo"), column(rows, "Name"));
    }

    @Test
    void aMergedNumberKeepsItsStoredValueAndItsDisplayedText() throws Exception {
        var rows = read(0, true);
        // C2:C3 holds 1234.5678 shown as 1,234.57 with the format #,##0.00; C4 is 99 and is not merged
        for (int i = 0; i < 2; i++) {
            assertEquals("1,234.57", rows.get(i).get("Amount").formattedValue(), "row " + (i + 2));
            assertEquals(0, new java.math.BigDecimal("1234.5678").compareTo(rows.get(i).get("Amount").rawNumber()), "row " + (i + 2));
        }
        assertEquals("99.00", rows.get(2).get("Amount").formattedValue());
        assertEquals(List.of("1,234.57", "1,234.57", "99.00", "10.00", "20.00", "5.00"),
                rows.stream().map(r -> r.get("Amount").formattedValue()).toList());
    }

    @Test
    void aRegionAcrossTwoColumnsFillsTheSecondColumnToo() throws Exception {
        // D2:E2 is one text over two columns; column E has no header, so it is read by position
        assertEquals("", read(0, false).get(0).get(4).asString());
        assertEquals("wide note", read(0, true).get(0).get(4).asString());
        assertEquals("wide note", read(0, true).get(0).get("Note").asString());
    }

    @Test
    void theSecondSheetHasItsOwnRegions() throws Exception {
        assertEquals(List.of("Other", ""), column(read(1, false), "Dept"));
        assertEquals(List.of("Other", "Other"), column(read(1, true), "Dept"));
        assertEquals(List.of("x", "y"), column(read(1, true), "Name"));
    }
}
