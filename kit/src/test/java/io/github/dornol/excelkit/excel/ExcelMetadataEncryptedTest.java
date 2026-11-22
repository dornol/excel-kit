package io.github.dornol.excelkit.excel;

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
        assertEquals(SHEETS, ExcelReader.getSheetNames(write(true), "secret"));
    }

    @Test
    void sheetNamesComeFromAStream() throws IOException {
        try (InputStream in = Files.newInputStream(write(true))) {
            assertEquals(SHEETS, ExcelReader.getSheetNames(in, "secret"));
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
        assertEquals(SHEETS, ExcelReader.getSheetNames(source, "secret"));
        assertEquals(1, opened[0]);
    }

    @Test
    void headersComeFromEachWayOfHandingOverTheFile() throws IOException {
        Path file = write(true);
        assertEquals(List.of("Name", "Age"), ExcelReader.getSheetHeaders(file, 0, 0, "secret"));
        assertEquals(List.of("City"), ExcelReader.getSheetHeaders(file, 1, 0, "secret"));
        try (InputStream in = Files.newInputStream(file)) {
            assertEquals(List.of("Name", "Age"), ExcelReader.getSheetHeaders(in, 0, 0, "secret"));
        }
        assertEquals(List.of("Name", "Age"), ExcelReader.getSheetHeaders(
                (InputStreamSource) () -> Files.newInputStream(file), 0, 0, "secret"));
    }

    @Test
    void aWrongPasswordIsReportedAsSuch() throws IOException {
        Path file = write(true);
        for (Runnable call : List.<Runnable>of(
                () -> ExcelReader.getSheetNames(file, "wrong"),
                () -> ExcelReader.getSheetHeaders(file, 0, 0, "wrong"))) {
            ExcelReadException error = assertThrows(ExcelReadException.class, call::run);
            assertTrue(error.getMessage().contains("Invalid password"), error.getMessage());
        }
    }

    @Test
    void aPasswordForAPlainWorkbookIsAFailureNotASilentRead() throws IOException {
        Path file = write(false);
        assertThrows(ExcelReadException.class, () -> ExcelReader.getSheetNames(file, "secret"));
        // the plain overloads still read it, and still cannot read an encrypted one
        assertEquals(SHEETS, ExcelReader.getSheetNames(file));
        assertThrows(ExcelReadException.class, () -> ExcelReader.getSheetNames(write(true)));
    }

    @Test
    void theHeaderPositionIsStillChecked() throws IOException {
        Path file = write(true);
        assertThrows(IllegalArgumentException.class, () -> ExcelReader.getSheetHeaders(file, -1, 0, "secret"));
        assertThrows(IllegalArgumentException.class, () -> ExcelReader.getSheetHeaders(file, 0, -1, "secret"));
    }
}
