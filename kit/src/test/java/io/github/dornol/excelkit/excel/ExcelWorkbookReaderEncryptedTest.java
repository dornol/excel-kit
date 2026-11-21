package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.RowError;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * A workbook that is encrypted as a whole is read sheet by sheet like any other, once each sheet's reader is
 * given the password. (The sheets cannot even be listed without decrypting the file first.)
 */
class ExcelWorkbookReaderEncryptedTest {

    @TempDir
    Path tempDir;

    public static class Item {
        Integer number;
    }

    private Path encryptedWorkbook() throws IOException {
        Path file = tempDir.resolve("encrypted.xlsx");
        try (ExcelWorkbook workbook = ExcelWorkbook.create().password("secret")) {
            workbook.<String>sheet("Numbers").column("number", s -> s).write(List.of("1", "x", "3", "y", "5"));
            workbook.<String>sheet("Other").column("number", s -> s).write(List.of("7", "8"));
            try (OutputStream out = Files.newOutputStream(file)) {
                workbook.finish().writeTo(out);
            }
        }
        return file;
    }

    private static ExcelReader<Item> reader(String password) {
        return ExcelReader.setter(Item::new).password(password).column("number", (item, cell) -> item.number = cell.asInt());
    }

    @Test
    void sheetsAreReadByNameAndByIndex() throws IOException {
        Path file = encryptedWorkbook();
        List<Integer> numbers = new ArrayList<>();
        List<Integer> other = new ArrayList<>();

        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet("Numbers", reader("secret"), result -> { if (result.success()) numbers.add(result.data().number); })
                .sheet(1, reader("secret"), result -> other.add(result.data().number))
                .read(file);

        assertEquals(List.of(1, 3, 5), numbers);
        assertEquals(List.of(7, 8), other);
        assertEquals("Numbers", report.sheets().get(0).name());
        assertEquals("Other", report.sheets().get(1).name());
        assertEquals(5, report.sheets().get(0).summary().totalRows());
        assertEquals(2, report.sheets().get(0).summary().errorRows());
        assertNull(report.sheets().get(1).failure());
    }

    @Test
    void anInputStreamIsReadToo() throws IOException {
        Path file = encryptedWorkbook();
        List<Integer> other = new ArrayList<>();
        try (var in = Files.newInputStream(file)) {
            ExcelWorkbookReader.create().sheet("Other", reader("secret"), result -> other.add(result.data().number)).read(in);
        }
        assertEquals(List.of(7, 8), other);
    }

    @Test
    void aWrongPasswordIsReportedAsSuch() throws IOException {
        Path file = encryptedWorkbook();
        ExcelReadException failure = assertThrows(ExcelReadException.class, () -> ExcelWorkbookReader.create()
                .sheet("Numbers", reader("wrong"), result -> { }).read(file));
        assertTrue(failure.getMessage().contains("Invalid password"), failure.getMessage());
    }

    @Test
    void errorsAreCollectedPerSheetUpToTheLimitWithTheirRowNumbers() throws IOException {
        Path file = encryptedWorkbook();
        List<Integer> seen = new ArrayList<>();

        WorkbookReadReport report = ExcelWorkbookReader.create().maxErrorsPerSheet(1)
                .sheet("Numbers", reader("secret"), result -> seen.add(result.success() ? result.data().number : null))
                .read(file);

        SheetReadReport sheet = report.sheets().get(0);
        assertEquals(5, seen.size(), "every row reaches the consumer, collected or not");
        assertEquals(1, sheet.errors().size());
        assertTrue(sheet.errorsTruncated());
        RowError first = sheet.errors().get(0);
        assertEquals(2, first.rowNum(), "the first failing row is the second data row");
        assertEquals(3, first.fileRowNum(), "which is row 3 of the sheet, below the header");
        assertNotNull(first.type());
    }

    @Test
    void allErrorsAreKeptWhenTheyFit() throws IOException {
        Path file = encryptedWorkbook();
        WorkbookReadReport report = ExcelWorkbookReader.create().maxErrorsPerSheet(2)
                .sheet("Numbers", reader("secret"), result -> { }).read(file);
        SheetReadReport sheet = report.sheets().get(0);
        assertEquals(2, sheet.errors().size());
        assertEquals(List.of(2L, 4L), sheet.errors().stream().map(RowError::rowNum).toList());
        assertTrue(!sheet.errorsTruncated());
    }

    @Test
    void aMissingSheetFollowsTheFailurePolicy() throws IOException {
        Path file = encryptedWorkbook();
        List<Integer> other = new ArrayList<>();

        WorkbookReadReport keepGoing = ExcelWorkbookReader.create()
                .sheet("Missing", reader("secret"), result -> { })
                .sheet("Other", reader("secret"), result -> other.add(result.data().number))
                .read(file);
        assertNotNull(keepGoing.sheets().get(0).failure());
        assertEquals(List.of(7, 8), other);

        other.clear();
        WorkbookReadReport stop = ExcelWorkbookReader.create().failurePolicy(SheetFailurePolicy.STOP)
                .sheet("Missing", reader("secret"), result -> { })
                .sheet("Other", reader("secret"), result -> other.add(result.data().number))
                .read(file);
        assertEquals(1, stop.sheets().size());
        assertTrue(other.isEmpty());
    }
}
