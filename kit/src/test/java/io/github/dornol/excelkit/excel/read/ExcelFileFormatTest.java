package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.excel.ExcelWriter;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.Consumer;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * A file that is not an .xlsx says what it is: an Excel 97-2003 workbook, a password-protected workbook read without its
 * password, or something else, from every way of reading (the readers, the inspector, the workbook reader).
 */
class ExcelFileFormatTest {

    @TempDir
    Path dir;

    private static byte[] xls() throws IOException {
        var out = new ByteArrayOutputStream();
        try (var wb = new HSSFWorkbook()) {
            var sheet = wb.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("Name");
            sheet.createRow(1).createCell(0).setCellValue("a");
            wb.write(out);
        }
        return out.toByteArray();
    }

    private static byte[] xlsx() throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().column("Name", s -> s).write(Stream.of("a", "b")).writeTo(out);
        return out.toByteArray();
    }

    private static byte[] encrypted(String password) throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().column("Name", s -> s).write(Stream.of("a")).writeTo(out, password);
        return out.toByteArray();
    }

    private static byte[] otherOle2() throws IOException {
        var out = new ByteArrayOutputStream();
        try (var fs = new POIFSFileSystem()) {
            fs.createDocument(new ByteArrayInputStream(new byte[16]), "WordDocument");
            fs.writeFilesystem(out);
        }
        return out.toByteArray();
    }

    private Path file(String name, byte[] bytes) throws IOException {
        Path path = dir.resolve(name);
        Files.write(path, bytes);
        return path;
    }

    /** Every way to read a workbook, each handed the bytes in the form it takes. */
    private List<Consumer<byte[]>> readers() {
        List<Consumer<byte[]>> all = new ArrayList<>();
        all.add(bytes -> ExcelReader.forMap().read(new ByteArrayInputStream(bytes), r -> {}));
        all.add(bytes -> ExcelReader.forMap().read(write("a.xlsx", bytes), r -> {}));
        all.add(bytes -> ExcelReader.<String>mapping(row -> row.get("Name").asString()).read(new ByteArrayInputStream(bytes), r -> {}));
        all.add(bytes -> ExcelReader.forMap().readStrict(new ByteArrayInputStream(bytes), r -> {}));
        all.add(bytes -> ExcelReader.forMap().readAll(new ByteArrayInputStream(bytes)));
        all.add(bytes -> ExcelReader.forMap().countRows().read(new ByteArrayInputStream(bytes), r -> {}));
        all.add(bytes -> ExcelWorkbookInspector.create().sheetNames(new ByteArrayInputStream(bytes)));
        all.add(bytes -> ExcelWorkbookInspector.create().sheetNames(write("b.xlsx", bytes)));
        all.add(bytes -> ExcelWorkbookInspector.create().headers(new ByteArrayInputStream(bytes), 0, 0));
        all.add(bytes -> ExcelWorkbookReader.create().sheet(0, ExcelReader.forMap(), r -> {}).read(new ByteArrayInputStream(bytes)));
        all.add(bytes -> ExcelWorkbookReader.create().sheet(0, ExcelReader.forMap(), r -> {}).read(write("c.xlsx", bytes)));
        return all;
    }

    private Path write(String name, byte[] bytes) {
        try {
            return file(name, bytes);
        } catch (IOException e) {
            throw new AssertionError(e);
        }
    }

    private void assertEveryWayExplains(byte[] bytes, String... expected) {
        int n = 0;
        for (var reader : readers()) {
            n++;
            var e = assertThrows(ExcelReadException.class, () -> reader.accept(bytes), "way " + n);
            for (String part : expected) {
                assertTrue(e.getMessage().contains(part), "way " + n + ": '" + e.getMessage() + "' says '" + part + "'");
            }
        }
    }

    @Nested
    class Xls {
        @Test
        void anXlsWorkbookSaysWhatItIsAndWhatToDoFromEveryWayOfReading() throws IOException {
            assertEveryWayExplains(xls(), "Excel 97-2003", ".xls", "save it as .xlsx");
        }

        @Test
        void theMessageIsOurOwnNotAWrappedPoiFailure() throws IOException {
            var e = assertThrows(ExcelReadException.class, () -> ExcelReader.forMap().read(new ByteArrayInputStream(xls()), r -> {}));
            assertTrue(e.getMessage().startsWith("The file is an Excel 97-2003 workbook"), e.getMessage());
            assertFalse(e.getMessage().contains("Failed to read"), "not wrapped");
        }

        @Test
        void aPasswordDoesNotMakeAnXlsReadable() throws IOException {
            var e = assertThrows(ExcelReadException.class,
                    () -> ExcelReader.forMap().password("pw").read(new ByteArrayInputStream(xls()), r -> {}));
            assertTrue(e.getMessage().contains(".xls"), e.getMessage());
        }

        @Test
        void anXlsWithTheNameOfAnXlsxIsStillRecognizedByItsContent() throws IOException {
            var e = assertThrows(ExcelReadException.class, () -> ExcelReader.forMap().read(file("report.xlsx", xls()), r -> {}));
            assertTrue(e.getMessage().contains("Excel 97-2003"));
        }

        @Test
        void anXlsxWithTheNameOfAnXlsIsStillRead() throws IOException {
            var rows = new ArrayList<String>();
            ExcelReader.forMap().read(file("report.xls", xlsx()), r -> rows.add(r.data().get("Name")));
            assertEquals(List.of("a", "b"), rows);
        }
    }

    @Nested
    class Encrypted {
        @Test
        void aPasswordProtectedWorkbookReadWithoutItsPasswordSaysSoFromEveryWayOfReading() throws IOException {
            assertEveryWayExplains(encrypted("pw"), "password-protected", "password(...)");
        }

        @Test
        void withTheRightPasswordItReadsAsBefore() throws IOException {
            var rows = new ArrayList<String>();
            ExcelReader.forMap().password("pw").read(new ByteArrayInputStream(encrypted("pw")), r -> rows.add(r.data().get("Name")));
            assertEquals(List.of("a"), rows);
            assertEquals(1, ExcelWorkbookInspector.create().password("pw").sheetNames(new ByteArrayInputStream(encrypted("pw"))).size());
        }

        @Test
        void aWrongPasswordIsStillReportedAsAWrongPassword() throws IOException {
            var e = assertThrows(ExcelReadException.class,
                    () -> ExcelReader.forMap().password("wrong").read(new ByteArrayInputStream(encrypted("pw")), r -> {}));
            assertTrue(e.getMessage().contains("Invalid password"), e.getMessage());
        }
    }

    @Nested
    class SomethingElse {
        @Test
        void anotherOle2DocumentIsNamedAsNotAWorkbook() throws IOException {
            assertEveryWayExplains(otherOle2(), "not an Excel workbook", "OLE2");
        }

        @Test
        void aTextOrCsvFileIsNamedAsNotAWorkbook() {
            assertEveryWayExplains("Name,Qty\na,1\n".getBytes(StandardCharsets.UTF_8), "not an Excel workbook", "ZIP archive", "CSV");
        }

        @Test
        void aPngIsNamedAsNotAWorkbook() {
            assertEveryWayExplains(new byte[]{(byte) 0x89, 'P', 'N', 'G', 0x0D, 0x0A, 0x1A, 0x0A, 0, 0, 0, 0}, "not an Excel workbook");
        }

        @Test
        void anEmptyFileIsNamedAsEmpty() {
            assertEveryWayExplains(new byte[0], "empty");
        }

        @Test
        void aFileOfOneByteIsNotAWorkbook() {
            assertEveryWayExplains(new byte[]{'x'}, "not an Excel workbook");
        }

        @Test
        void aZipThatIsNoWorkbookKeepsTheReportOfTheReadItself() throws IOException {
            var out = new ByteArrayOutputStream();
            try (var zip = new java.util.zip.ZipOutputStream(out)) {
                zip.putNextEntry(new java.util.zip.ZipEntry("readme.txt"));
                zip.write("hello".getBytes(StandardCharsets.UTF_8));
                zip.closeEntry();
            }
            var e = assertThrows(ExcelKitException.class, () -> ExcelReader.forMap().read(new ByteArrayInputStream(out.toByteArray()), r -> {}));
            assertFalse(String.valueOf(e.getMessage()).startsWith("The file is"), "a ZIP is not judged here: " + e.getMessage());
        }
    }

    @Nested
    class Unchanged {
        @Test
        void anXlsxReadsAsBeforeFromEveryWayOfReading() throws IOException {
            byte[] bytes = xlsx();
            var rows = new ArrayList<String>();
            ExcelReader.forMap().read(new ByteArrayInputStream(bytes), r -> rows.add(r.data().get("Name")));
            ExcelReader.forMap().read(write("ok.xlsx", bytes), r -> rows.add(r.data().get("Name")));
            assertEquals(List.of("a", "b", "a", "b"), rows);
            assertEquals(1, ExcelWorkbookInspector.create().sheetNames(new ByteArrayInputStream(bytes)).size());
            assertEquals(List.of("Name"), ExcelWorkbookInspector.create().headers(new ByteArrayInputStream(bytes), 0, 0));
            var report = ExcelWorkbookReader.create().sheet(0, ExcelReader.forMap(), r -> {}).read(new ByteArrayInputStream(bytes));
            assertNotNull(report);
        }

        @Test
        void aFileThatCannotBeOpenedIsLeftToTheReadItself() {
            assertThrows(ExcelKitException.class, () -> ExcelReader.forMap().read(dir.resolve("missing.xlsx"), r -> {}));
        }
    }
}
