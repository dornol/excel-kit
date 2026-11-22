package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.io.InputStreamSource;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The ways of registering sheets with an {@link ExcelWorkbookReader}, and what it does when a sheet or the input is not usable. */
class ExcelWorkbookReaderPathsTest {

    private static byte[] workbook() {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("First").column("A", v -> v).write(List.of("a1", "a2").stream());
            workbook.<String>sheet("Second").column("B", v -> v).write(List.of("b1").stream());
            workbook.<String>sheet("Third").column("C", v -> v).write(List.of("c1", "c2", "c3").stream());
            workbook.finish().writeTo(output);
        }
        return output.toByteArray();
    }

    private static ExcelReader<String> firstColumn() {
        return ExcelReader.mapping(row -> row.get(0).asString());
    }

    // ---- sheets by index ----

    @Test
    void sheetsRegisteredByIndexAreReadAndReportedUnderTheirNames() {
        List<String> second = new ArrayList<>();
        List<String> third = new ArrayList<>();

        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet(1, firstColumn(), r -> second.add(r.data()))
                .sheet(2, firstColumn(), r -> third.add(r.data()))
                .read(new ByteArrayInputStream(workbook()));

        assertTrue(report.successful());
        assertEquals(List.of("b1"), second);
        assertEquals(List.of("c1", "c2", "c3"), third);
        assertEquals(List.of("Second", "Third"), report.sheets().stream().map(SheetReadReport::name).toList());
        assertEquals(List.of(1, 2), report.sheets().stream().map(SheetReadReport::index).toList());
    }

    @Test
    void aSheetByIndexRefusesBadArguments() {
        var reader = ExcelWorkbookReader.create();

        assertThrows(IllegalArgumentException.class, () -> reader.sheet(-1, firstColumn(), r -> { }));
        assertThrows(NullPointerException.class, () -> reader.sheet(0, null, r -> { }));
        assertThrows(NullPointerException.class, () -> reader.sheet(0, firstColumn(), null));
    }

    @Test
    void anIndexBeyondTheLastSheetIsAFailureOfThatSheetAndTheOthersStillRead() {
        List<String> first = new ArrayList<>();

        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet(9, firstColumn(), r -> { })
                .sheet(0, firstColumn(), r -> first.add(r.data()))
                .read(new ByteArrayInputStream(workbook()));

        assertFalse(report.successful());
        assertEquals(2, report.sheets().size());
        assertNotNull(report.sheets().get(0).failure());
        assertEquals("Sheet index 9 is not present", report.sheets().get(0).failure().getMessage());
        assertEquals(List.of("a1", "a2"), first);
    }

    @Test
    void stopPolicyLeavesTheSheetsAfterABadIndexUnread() {
        List<String> first = new ArrayList<>();

        WorkbookReadReport report = ExcelWorkbookReader.create().failurePolicy(SheetFailurePolicy.STOP)
                .sheet(9, firstColumn(), r -> { })
                .sheet(0, firstColumn(), r -> first.add(r.data()))
                .read(new ByteArrayInputStream(workbook()));

        assertEquals(1, report.sheets().size());
        assertTrue(first.isEmpty());
    }

    // ---- error limits ----

    @Test
    void theNumberOfKeptErrorsIsChosenPerSheet() {
        // 'B' holds text, so converting it to a number fails for every row
        ExcelReader<Integer> failing = ExcelReader.mapping(row -> row.get(0).asInt());

        WorkbookReadReport none = ExcelWorkbookReader.create().maxErrorsPerSheet(0)
                .sheet("Third", failing, r -> { }).read(new ByteArrayInputStream(workbook()));
        WorkbookReadReport two = ExcelWorkbookReader.create().maxErrorsPerSheet(2)
                .sheet("Third", failing, r -> { }).read(new ByteArrayInputStream(workbook()));

        assertEquals(3, none.sheet("Third").summary().errorRows());
        assertTrue(none.sheet("Third").errors().isEmpty());
        assertTrue(none.sheet("Third").errorsTruncated());
        assertEquals(2, two.sheet("Third").errors().size());
        assertTrue(two.sheet("Third").errorsTruncated());
    }

    @Test
    void aNegativeErrorLimitIsRefused() {
        assertThrows(IllegalArgumentException.class, () -> ExcelWorkbookReader.create().maxErrorsPerSheet(-1));
    }

    // ---- inputs that fail ----

    private static final class BrokenInput extends InputStream {
        @Override
        public int read() throws IOException {
            throw new IOException("connection reset");
        }
    }

    @Test
    void aStreamThatFailsWhileItIsCopiedIsReportedAsALibraryException() {
        var reader = ExcelWorkbookReader.create().sheet(0, firstColumn(), r -> { });

        ExcelKitException e = assertThrows(ExcelKitException.class, () -> reader.read(new BrokenInput()));

        assertNotNull(e.getCause(), String.valueOf(e));
    }

    @Test
    void aSourceThatCannotBeOpenedIsReportedAsALibraryException() {
        InputStreamSource source = () -> {
            throw new IOException("no such object");
        };
        var reader = ExcelWorkbookReader.create().sheet(0, firstColumn(), r -> { });

        ExcelKitException e = assertThrows(ExcelKitException.class, () -> reader.read(source));

        assertEquals("Failed to open workbook input", e.getMessage());
        assertEquals("no such object", e.getCause().getMessage());
    }

    @Test
    void aSourceIsOpenedAndClosedByTheRead() throws Exception {
        boolean[] closed = {false};
        byte[] bytes = workbook();
        InputStreamSource source = () -> new ByteArrayInputStream(bytes) {
            @Override
            public void close() throws IOException {
                closed[0] = true;
                super.close();
            }
        };

        WorkbookReadReport report = ExcelWorkbookReader.create().sheet(0, firstColumn(), r -> { }).read(source);

        assertTrue(report.successful());
        assertTrue(closed[0], "the stream the source opened is closed");
    }

    // ---- passwords and nothing to read ----

    @Test
    void aPasswordGivenForAWorkbookThatIsNotEncryptedFailsTheOpening(@TempDir Path dir) throws Exception {
        Path file = dir.resolve("book.xlsx");
        Files.write(file, workbook());

        ExcelKitException e = assertThrows(ExcelKitException.class, () -> ExcelWorkbookReader.create()
                .sheet(0, firstColumn().password("unused"), r -> { }).read(file));

        assertNotNull(e.getMessage());
    }

    @Test
    void withoutAnyRegisteredSheetNothingIsReadAndTheReportIsEmpty() {
        WorkbookReadReport report = ExcelWorkbookReader.create().read(new ByteArrayInputStream(workbook()));

        assertTrue(report.sheets().isEmpty());
        assertTrue(report.successful());
    }

    @Test
    void anEncryptedWorkbookIsDecryptedOnceForEverySheet(@TempDir Path dir) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook workbook = ExcelWorkbook.create().password("pw")) {
            workbook.<String>sheet("First").column("A", v -> v).write(List.of("a1").stream());
            workbook.<String>sheet("Second").column("B", v -> v).write(List.of("b1").stream());
            workbook.finish().writeTo(out);
        }
        Path file = dir.resolve("secret.xlsx");
        Files.write(file, out.toByteArray());
        List<String> rows = new ArrayList<>();

        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet(0, firstColumn().password("pw"), r -> rows.add(r.data()))
                .sheet("Second", firstColumn().password("pw"), r -> rows.add(r.data()))
                .read(file);

        assertTrue(report.successful(), String.valueOf(report.sheets().stream().map(SheetReadReport::failure).toList()));
        assertEquals(List.of("a1", "b1"), rows);
    }
}
