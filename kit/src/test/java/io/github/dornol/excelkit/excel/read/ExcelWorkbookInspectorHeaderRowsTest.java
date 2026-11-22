package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;

/**
 * {@code headers(input, sheetIndex, headerRowIndex, headerRows)} names the columns of a workbook whose header spans
 * several rows, as the reader's {@code headerRows} does: the block ends at {@code headerRowIndex}, and for each column
 * the lowest non-blank cell of the block is its name. Without it no single row of a workbook written with
 * {@code group(...)} holds every column name.
 */
class ExcelWorkbookInspectorHeaderRowsTest {

    private record Row(String a, String b, String c, String d) {
    }

    private static List<Row> rows(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Row("a" + i, "b" + i, "c" + i, "d" + i)).toList();
    }

    /** First | (G: Grouped1, Grouped2) | No. : the last column has no group and is merged down over both header rows. */
    private static byte[] grouped() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().column("First", Row::a).column("Grouped1", Row::b, c -> c.group("G"))
                .column("Grouped2", Row::c, c -> c.group("G")).rowNumberColumn("No.")
                .write(rows(3)).writeTo(out);
        return out.toByteArray();
    }

    private static List<String> headers(byte[] xlsx, int sheet, int headerRowIndex, int headerRows) {
        return ExcelWorkbookInspector.create().headers(new ByteArrayInputStream(xlsx), sheet, headerRowIndex, headerRows);
    }

    @Test
    void aGroupedHeaderGivesEveryColumnName() {
        assertEquals(List.of("First", "Grouped1", "Grouped2", "No."), headers(grouped(), 0, 1, 2));
    }

    @Test
    void itAgreesWithWhatTheReaderCallsTheColumns() {
        byte[] xlsx = grouped();
        List<String> readerNames = new ArrayList<>();
        ExcelReader.forMap().headerRowIndex(1).headerRows(2).read(new ByteArrayInputStream(xlsx),
                r -> { if (readerNames.isEmpty()) readerNames.addAll(r.data().keySet()); });

        assertEquals(readerNames, headers(xlsx, 0, 1, 2));
    }

    @Test
    void severalLevelsOfGroupsAreCombinedToo() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create()
                .column("X", Row::a, c -> c.group("Top", "Mid"))
                .column("Y", Row::b, c -> c.group("Top"))
                .column("Z", Row::c)
                .write(rows(2)).writeTo(out);

        assertEquals(List.of("X", "Y", "Z"), headers(out.toByteArray(), 0, 2, 3));
    }

    @Test
    void oneHeaderRowIsReadAsBefore() {
        byte[] xlsx = grouped();
        var inspector = ExcelWorkbookInspector.create();

        assertEquals(inspector.headers(new ByteArrayInputStream(xlsx), 0, 1), headers(xlsx, 0, 1, 1));
        assertEquals(List.of("", "Grouped1", "Grouped2"), headers(xlsx, 0, 1, 1));
        assertEquals(List.of("First", "G", "G", "No."), headers(xlsx, 0, 0, 1));
    }

    @Test
    void aBlockThatStartsAboveTheSheetIsClippedLikeTheReaderDoes() {
        // headerRowIndex 1 with 5 rows asks for rows -3..1: only rows 0 and 1 exist
        assertEquals(List.of("First", "Grouped1", "Grouped2", "No."), headers(grouped(), 0, 1, 5));
    }

    @Test
    void everySheetOfARolledOverWorkbookHasTheSameNames() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().maxRows(2).column("First", Row::a).column("Grouped1", Row::b, c -> c.group("G"))
                .rowNumberColumn("No.").write(rows(5)).writeTo(out);

        for (int sheet = 0; sheet < 3; sheet++) {
            assertEquals(List.of("First", "Grouped1", "No."), headers(out.toByteArray(), sheet, 1, 2), "sheet " + sheet);
        }
    }

    @Test
    void aPathAndAnInputStreamSourceGiveTheSame(@TempDir Path dir) throws Exception {
        byte[] xlsx = grouped();
        Path file = dir.resolve("grouped.xlsx");
        Files.write(file, xlsx);
        InputStreamSource source = () -> new ByteArrayInputStream(xlsx);
        var inspector = ExcelWorkbookInspector.create();

        List<String> expected = List.of("First", "Grouped1", "Grouped2", "No.");
        assertEquals(expected, inspector.headers(file, 0, 1, 2));
        assertEquals(expected, inspector.headers(source, 0, 1, 2));
    }

    @Test
    void anEncryptedWorkbookIsOpenedWithItsPassword() {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Row>create().password("pw").column("First", Row::a).column("Grouped1", Row::b, c -> c.group("G"))
                .rowNumberColumn("No.").write(rows(2)).writeTo(out);

        assertEquals(List.of("First", "Grouped1", "No."),
                ExcelWorkbookInspector.create().password("pw").headers(new ByteArrayInputStream(out.toByteArray()), 0, 1, 2));
    }

    @Test
    void aHeaderBlockOfNoRowsIsRefused() {
        byte[] xlsx = grouped();
        var inspector = ExcelWorkbookInspector.create();

        assertThrows(IllegalArgumentException.class, () -> inspector.headers(new ByteArrayInputStream(xlsx), 0, 1, 0));
        assertThrows(IllegalArgumentException.class, () -> inspector.headers(new ByteArrayInputStream(xlsx), 0, 1, -1));
        assertThrows(IllegalArgumentException.class, () -> inspector.headers(new ByteArrayInputStream(xlsx), 0, -1, 1));
        assertThrows(IllegalArgumentException.class, () -> inspector.headers(new ByteArrayInputStream(xlsx), -1, 1, 1));
    }
}
