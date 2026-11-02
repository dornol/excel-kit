package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.excel.ExcelReadException;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import java.io.ByteArrayInputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Objects;
import java.util.concurrent.atomic.AtomicReference;
import static org.junit.jupiter.api.Assertions.*;

class ExternalEncryptedWorkbookTest {
    private static final String FIXTURE_PASSWORD = "Test001!!"; // ggignore: public Apache POI fixture credential; provenance in external-workbooks/README.md.

    @Test
    void macExcelEncryptedWorkbookReadsByPathStreamAndSourceWithoutChangingTheFixture() throws Exception {
        Path file = fixture();
        byte[] original = Files.readAllBytes(file);
        for (int inputKind = 0; inputKind < 3; inputKind++) {
            var rows = new ArrayList<ReadResult<RowData>>();
            var reader = ExcelReader.<RowData>mapping(row -> row).skipBlankRows().password(FIXTURE_PASSWORD);
            var input = new TrackedInput(original);
            try (var session = inputKind == 0 ? reader.open(file)
                    : inputKind == 1 ? reader.open(input) : reader.open((InputStreamSource) () -> input)) {
                var container = (TempResourceContainer) session;
                var directory = new AtomicReference<Path>();
                session.read(row -> {
                    assertTrue(row.success());
                    directory.set(container.getTempDir());
                    assertTrue(Files.isDirectory(directory.get()));
                    rows.add(row);
                });
                assertNotNull(directory.get());
                assertFalse(Files.exists(directory.get()), "Decryption workspace must be removed");
            }
            assertEquals(4, rows.size());
            assertEquals("This WB was created with Excel for Mac 2011", rows.get(2).data().get(1).asString());
            assertEquals("Reencrypted by Mac Excel 2016", rows.get(3).data().get(1).asString());
            assertEquals(inputKind == 2, input.closed);
            assertArrayEquals(original, Files.readAllBytes(file));
        }
        try (var input = Files.newInputStream(file);
             var workbook = (XSSFWorkbook) WorkbookFactory.create(input, FIXTURE_PASSWORD)) {
            var properties = workbook.getProperties().getExtendedProperties().getUnderlyingProperties();
            assertEquals("Microsoft Macintosh Excel", properties.getApplication());
            assertEquals("15.0300", properties.getAppVersion());
        }
    }

    @Test
    void decryptionWorkspaceIsRemovedWhenPreflightSecurityRejectsTheWorkbook() throws Exception {
        Path file = fixture();
        byte[] original = Files.readAllBytes(file);
        for (boolean pathInput : List.of(false, true)) {
            var reader = ExcelReader.forMap().password(FIXTURE_PASSWORD)
                    .securityPolicy(new ReadSecurityPolicy(true, true, 1_000_000, 2_000_000, 100, 1));
            try (var session = pathInput ? reader.open(file) : reader.open(new ByteArrayInputStream(original))) {
                var container = (TempResourceContainer) session;
                var error = assertThrows(ReadSecurityException.class, () -> session.read(row -> fail("No rows before validation")));
                assertEquals(ReadSecurityException.Reason.ZIP_ENTRIES, error.reason());
                assertNotNull(container.getTempDir());
                assertFalse(Files.exists(container.getTempDir()), "Rejected plaintext and its directory must be removed");
            }
            assertArrayEquals(original, Files.readAllBytes(file));
        }
    }

    @Test
    void wrongPasswordDoesNotDeleteOrRewriteTheEncryptedSource() throws Exception {
        Path file = fixture();
        byte[] original = Files.readAllBytes(file);
        var error = assertThrows(ExcelReadException.class, () -> ExcelReader.forMap()
                .password("incorrect").read(file, row -> fail()));
        assertTrue(error.getMessage().contains("password"));
        assertArrayEquals(original, Files.readAllBytes(file));
    }

    private static Path fixture() throws Exception {
        return Path.of(Objects.requireNonNull(ExternalEncryptedWorkbookTest.class
                .getResource("/external-workbooks/60320-protected.xlsx")).toURI());
    }

    private static final class TrackedInput extends ByteArrayInputStream {
        boolean closed;
        TrackedInput(byte[] data) { super(data); }
        @Override public void close() { closed = true; }
    }
}
