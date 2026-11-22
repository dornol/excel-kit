package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.read.ReadSecurityException;
import io.github.dornol.excelkit.core.read.ReadSecurityPolicy;
import io.github.dornol.excelkit.excel.FeatureMatrixTest.Feature;
import io.github.dornol.excelkit.excel.FeatureMatrixTest.Item;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import io.github.dornol.excelkit.excel.read.ExcelWorkbookInspector;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static io.github.dornol.excelkit.excel.FeatureMatrixTest.describe;
import static io.github.dornol.excelkit.excel.FeatureMatrixTest.on;
import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * What {@link FeatureMatrixTest} writes, read back with {@link ExcelReader}: every combination of ten sheet features
 * must give its items, in order, with their numbers, across the sheets it rolled over to, whatever sits around them
 * (group headers that need a second header row, summary rows, a table, protection, an encrypted file). The same file
 * is also described by {@link ExcelWorkbookInspector}, and read under the strict security policy, which must refuse
 * exactly the files that hold formulas (the summary rows).
 */
class FeatureMatrixReadBackTest {

    private static final int ROWS = FeatureMatrixTest.ROWS;

    private static IntStream masks() {
        return IntStream.range(0, 1 << Feature.values().length);
    }

    private record Seen(String name, String qty, String number) {
    }

    private static List<Seen> readSheet(byte[] xlsx, int sheet, int headerRows, boolean rowNumbers, boolean encrypted,
                                        ReadSecurityPolicy policy) {
        List<Seen> seen = new ArrayList<>();
        ExcelReader<Seen> reader = ExcelReader.<Seen>mapping(row -> new Seen(row.get("Name").asString(), row.get("Qty").asString(),
                rowNumbers ? row.get("No.").asString() : "")).sheetIndex(sheet).headerRowIndex(headerRows - 1).headerRows(headerRows);
        if (encrypted) reader.password("pw");
        if (policy != null) reader.securityPolicy(policy);
        reader.read(new ByteArrayInputStream(xlsx), (ReadResult<Seen> r) -> {
            assertTrue(r.success(), "row " + r.fileRowNum() + ": " + r.messages());
            seen.add(r.data());
        });
        return seen;
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("masks")
    void everyCombinationReadsBackWithTheReader(int mask) throws Exception {
        String at = describe(mask);
        boolean roll = on(mask, Feature.ROLLOVER), group = on(mask, Feature.GROUP), summary = on(mask, Feature.SUMMARY);
        boolean rowNumbers = on(mask, Feature.ROW_NUMBER), encrypted = on(mask, Feature.ENCRYPT);
        int perSheet = roll ? 4 : ROWS;
        int sheets = roll ? (int) Math.ceil(ROWS / (double) perSheet) : 1;
        int headerRows = group ? 2 : 1;

        List<Item> items = IntStream.range(0, ROWS).mapToObj(i -> new Item("item " + i, i)).toList();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        FeatureMatrixTest.buildWriter(mask).write(items).writeTo(out);
        byte[] xlsx = out.toByteArray();

        // ---- the reader ----
        int global = 0;
        for (int s = 0; s < sheets; s++) {
            List<Seen> seen = readSheet(xlsx, s, headerRows, rowNumbers, encrypted, null);
            int expected = Math.min(perSheet, ROWS - global);
            List<Seen> data = seen.stream().filter(x -> x.name().startsWith("item ")).toList();
            List<Seen> others = seen.stream().filter(x -> !x.name().startsWith("item ")).toList();
            assertEquals(expected, data.size(), at + " items on sheet " + s);
            for (int i = 0; i < data.size(); i++, global++) {
                assertEquals("item " + global, data.get(i).name(), at + " name of item " + global);
                assertEquals(String.valueOf(global), data.get(i).qty(), at + " qty of item " + global);
                if (rowNumbers) assertEquals(String.valueOf(global + 1), data.get(i).number(), at + " row number of item " + global);
            }
            // the only other row is the summary, which comes after the data
            assertEquals(summary ? List.of("Total") : List.of(), others.stream().map(Seen::name).toList(), at + " other rows on sheet " + s);
            if (summary) assertEquals("Total", seen.get(seen.size() - 1).name(), at + " the summary is last on sheet " + s);
        }
        assertEquals(ROWS, global, at);

        // ---- the same file described without reading rows ----
        var inspector = ExcelWorkbookInspector.create();
        if (encrypted) inspector.password("pw");
        assertEquals(sheets, inspector.sheetNames(new ByteArrayInputStream(xlsx)).size(), at + " sheets");
        List<String> headers = inspector.headers(new ByteArrayInputStream(xlsx), 0, headerRows - 1);
        List<String> expectedHeaders = new ArrayList<>(List.of("Name", "Qty"));
        // With group headers the name of a column without a group sits in the first header row (the cell is merged
        // down), which the inspector, reading one row, does not see; see inspectorReadsOneRowOfAGroupedHeader.
        if (rowNumbers && !group) expectedHeaders.add("No.");
        assertEquals(expectedHeaders, headers, at + " headers");

        // ---- the strict policy refuses formulas, and nothing else here ----
        if (summary) {
            assertThrows(ReadSecurityException.class,
                    () -> readSheet(xlsx, 0, headerRows, rowNumbers, encrypted, ReadSecurityPolicy.STRICT), at + " strict policy");
        } else {
            assertEquals(readSheet(xlsx, 0, headerRows, rowNumbers, encrypted, null),
                    readSheet(xlsx, 0, headerRows, rowNumbers, encrypted, ReadSecurityPolicy.STRICT), at + " strict policy");
        }
    }

    private record Wide(String a, String b, String c) {
    }

    /**
     * Pins what {@link ExcelWorkbookInspector#headers} does with a grouped header today: it returns the one row it is
     * asked for. The reader combines the rows of the header block ({@code headerRows}); the inspector has no such option,
     * so for a workbook written with {@code group(...)} no single row holds every column name.
     */
    @org.junit.jupiter.api.Test
    void inspectorReadsOneRowOfAGroupedHeader() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Wide>create().column("First", Wide::a).column("Grouped1", Wide::b, c -> c.group("G"))
                .column("Grouped2", Wide::c, c -> c.group("G")).rowNumberColumn("No.")
                .write(List.of(new Wide("x", "y", "z"))).writeTo(out);
        var inspector = ExcelWorkbookInspector.create();

        assertEquals(List.of("First", "G", "G", "No."), inspector.headers(new ByteArrayInputStream(out.toByteArray()), 0, 0));
        assertEquals(List.of("", "Grouped1", "Grouped2"), inspector.headers(new ByteArrayInputStream(out.toByteArray()), 0, 1));

        // the reader, given the header block, does know every name
        List<String> names = new ArrayList<>();
        ExcelReader.forMap().headerRowIndex(1).headerRows(2).read(new ByteArrayInputStream(out.toByteArray()),
                r -> names.addAll(r.data().keySet()));
        assertEquals(List.of("First", "Grouped1", "Grouped2", "No."), names);
    }
}
