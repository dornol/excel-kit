package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.excel.ExcelWorkbook;
import io.github.dornol.excelkit.core.io.InputStreamSource;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** Sheet names and headers of a password-protected workbook, through each way of handing it over. */
class ExcelMetadataEncryptedTest {

    @TempDir
    Path tempDir;

    private Path write(boolean encrypted) throws IOException {
        Path file = tempDir.resolve(encrypted ? "encrypted.xlsx" : "plain.xlsx");
        try (ExcelWorkbook workbook = encrypted ? ExcelWorkbook.create().password("secret") : ExcelWorkbook.create()) {
            workbook.<String>sheet("First").column("Name", s -> s).column("Age", s -> s).write(List.of("a"));
            workbook.<String>sheet("Second").column("City", s -> s).write(List.of("b"));
            try (OutputStream out = Files.newOutputStream(file)) {
                workbook.finish().writeTo(out);
            }
        }
        return file;
    }

    private static final List<ExcelSheetInfo> SHEETS = List.of(new ExcelSheetInfo(0, "First"), new ExcelSheetInfo(1, "Second"));

    @Test
    void sheetNamesComeFromAPath() throws IOException {
        assertEquals(SHEETS, ExcelWorkbookInspector.create().password("secret").sheetNames(write(true)));
    }

    @Test
    void sheetNamesComeFromAStream() throws IOException {
        try (InputStream in = Files.newInputStream(write(true))) {
            assertEquals(SHEETS, ExcelWorkbookInspector.create().password("secret").sheetNames(in));
        }
    }

    @Test
    void sheetNamesComeFromASourceThatIsOpenedAndClosed() throws IOException {
        Path file = write(true);
        int[] opened = {0};
        InputStreamSource source = () -> {
            opened[0]++;
            return Files.newInputStream(file);
        };
        assertEquals(SHEETS, ExcelWorkbookInspector.create().password("secret").sheetNames(source));
        assertEquals(1, opened[0]);
    }

    @Test
    void headersComeFromEachWayOfHandingOverTheFile() throws IOException {
        Path file = write(true);
        assertEquals(List.of("Name", "Age"), ExcelWorkbookInspector.create().password("secret").headers(file, 0, 0));
        assertEquals(List.of("City"), ExcelWorkbookInspector.create().password("secret").headers(file, 1, 0));
        try (InputStream in = Files.newInputStream(file)) {
            assertEquals(List.of("Name", "Age"), ExcelWorkbookInspector.create().password("secret").headers(in, 0, 0));
        }
        assertEquals(List.of("Name", "Age"), ExcelWorkbookInspector.create().password("secret").headers(
                (InputStreamSource) () -> Files.newInputStream(file), 0, 0));
    }

    @Test
    void aWrongPasswordIsReportedAsSuch() throws IOException {
        Path file = write(true);
        for (Runnable call : List.<Runnable>of(
                () -> ExcelWorkbookInspector.create().password("wrong").sheetNames(file),
                () -> ExcelWorkbookInspector.create().password("wrong").headers(file, 0, 0))) {
            ExcelReadException error = assertThrows(ExcelReadException.class, call::run);
            assertTrue(error.getMessage().contains("Invalid password"), error.getMessage());
        }
    }

    @Test
    void aPasswordForAPlainWorkbookIsAFailureNotASilentRead() throws IOException {
        Path file = write(false);
        assertThrows(ExcelReadException.class, () -> ExcelWorkbookInspector.create().password("secret").sheetNames(file));
        // the plain overloads still read it, and still cannot read an encrypted one
        assertEquals(SHEETS, ExcelWorkbookInspector.create().sheetNames(file));
        assertThrows(ExcelReadException.class, () -> ExcelWorkbookInspector.create().sheetNames(write(true)));
    }

    @Test
    void theHeaderPositionIsStillChecked() throws IOException {
        Path file = write(true);
        assertThrows(IllegalArgumentException.class, () -> ExcelWorkbookInspector.create().password("secret").headers(file, -1, 0));
        assertThrows(IllegalArgumentException.class, () -> ExcelWorkbookInspector.create().password("secret").headers(file, 0, -1));
    }
}
