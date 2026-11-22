package io.github.dornol.excelkit.spring;

import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.Arrays;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** What {@link ExcelKitMultipartFile} refuses, and what it says about it. */
class ExcelKitMultipartFileFailuresTest {

    private static String refusal(Runnable validation) {
        return assertThrows(ExcelKitUploadException.class, validation::run).getMessage();
    }

    private static byte[] xlsxLikeZip() throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ZipOutputStream zip = new ZipOutputStream(out)) {
            zip.putNextEntry(new ZipEntry("[Content_Types].xml"));
            zip.write("<Types/>".getBytes(StandardCharsets.UTF_8));
            zip.closeEntry();
            zip.putNextEntry(new ZipEntry("xl/workbook.xml"));
            zip.write("<workbook/>".getBytes(StandardCharsets.UTF_8));
            zip.closeEntry();
        }
        return out.toByteArray();
    }

    @Nested
    class Opening {
        @Test
        void aStreamThatCannotBeOpenedIsReportedWithTheFilenameAndItsCause() {
            ExcelKitUploadException e = assertThrows(ExcelKitUploadException.class,
                    () -> ExcelKitMultipartFile.open(StubMultipartFile.unopenable("report.xlsx")));

            assertEquals("Failed to open upload file: report.xlsx", e.getMessage());
            assertEquals("disk gone", e.getCause().getMessage());
        }

        @Test
        void theFilenameInAMessageHasNoControlCharactersOrPathSeparators() {
            ExcelKitUploadException e = assertThrows(ExcelKitUploadException.class,
                    () -> ExcelKitMultipartFile.open(StubMultipartFile.unopenable("a\nb/c\\d.xlsx")));

            assertEquals("Failed to open upload file: a_b_c_d.xlsx", e.getMessage());
        }

        @Test
        void aMissingOrBlankFilenameIsShownAsUnknown() {
            assertEquals("Failed to open upload file: <unknown>", assertThrows(ExcelKitUploadException.class,
                    () -> ExcelKitMultipartFile.open(StubMultipartFile.unopenable(null))).getMessage());
            assertEquals("Failed to open upload file: <unknown>", assertThrows(ExcelKitUploadException.class,
                    () -> ExcelKitMultipartFile.open(StubMultipartFile.unopenable("  "))).getMessage());
        }

        @Test
        void aNullFileIsRefused() {
            assertThrows(NullPointerException.class, () -> ExcelKitMultipartFile.open(null));
        }
    }

    @Nested
    class Size {
        @Test
        void aFileAboveTheLimitIsRefusedAndOneAtItIsNot() {
            MultipartFile file = new MockMultipartFile("file", "a.csv", "text/csv", new byte[10]);

            assertEquals("Upload file exceeds maximum size of 9 bytes", refusal(() -> ExcelKitMultipartFile.requireSizeAtMost(file, 9)));
            assertSame(file, ExcelKitMultipartFile.requireSizeAtMost(file, 10));
        }

        @Test
        void aNegativeLimitIsRefused() {
            MultipartFile file = new MockMultipartFile("file", "a.csv", "text/csv", new byte[1]);

            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireSizeAtMost(file, -1));
        }
    }

    @Nested
    class Extensions {
        private MultipartFile named(String filename) {
            return new MockMultipartFile("file", filename, "application/octet-stream", new byte[]{1});
        }

        @Test
        void theRefusalListsTheAllowedExtensionsSortedAndWithoutDots() {
            assertEquals("Unsupported upload file extension. Allowed extensions: csv, xlsx",
                    refusal(() -> ExcelKitMultipartFile.requireExtension(named("a.pdf"), ".XLSX", "csv")));
        }

        @Test
        void aFilenameWithoutARealExtensionHasNone() {
            for (String filename : Arrays.asList("noextension", "trailingdot.", "dir.d/file", "dir.d\\file", ".hidden", "")) {
                assertEquals("Unsupported upload file extension. Allowed extensions: xlsx",
                        refusal(() -> ExcelKitMultipartFile.requireExtension(named(filename), "xlsx")), "<" + filename + ">");
            }
        }

        @Test
        void theExtensionIsCaseInsensitive() {
            MultipartFile file = named("REPORT.XLSX");

            assertSame(file, ExcelKitMultipartFile.requireExtension(file, "xlsx"));
        }

        @Test
        void aMissingFilenameHasNoExtension() {
            // MockMultipartFile turns a null filename into "", so a file that really returns null is a stub
            MultipartFile file = StubMultipartFile.of(null, new byte[]{1});

            assertEquals("Unsupported upload file extension. Allowed extensions: xlsx",
                    refusal(() -> ExcelKitMultipartFile.requireExtension(file, "xlsx")));
        }

        @Test
        void theAllowedExtensionsMustBeGivenAndBeSimple() {
            MultipartFile file = named("a.xlsx");

            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireExtension(file));
            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireExtension(file, (String[]) null));
            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireExtension(file, " "));
            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireExtension(file, (String) null));
            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireExtension(file, "..."));
            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireExtension(file, "a/b"));
            assertThrows(IllegalArgumentException.class, () -> ExcelKitMultipartFile.requireExtension(file, "a\\b"));
        }
    }

    @Nested
    class Content {
        @Test
        void aTextFileCalledXlsxIsNotAWorkbook() {
            MultipartFile file = new MockMultipartFile("file", "fake.xlsx", "text/plain", "a,b\n1,2\n".getBytes(StandardCharsets.UTF_8));

            assertEquals("Upload file content is not a valid XLSX/XLSM package",
                    refusal(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(file)));
        }

        @Test
        void aZipPackageIsAcceptedAsXlsxAndXlsm() throws Exception {
            for (String name : new String[]{"book.xlsx", "book.xlsm"}) {
                MultipartFile file = new MockMultipartFile("file", name, "application/octet-stream", xlsxLikeZip());

                assertSame(file, ExcelKitMultipartFile.requireExcelOrCsvContent(file), name);
            }
        }

        @Test
        void aBinaryFileCalledCsvIsNotText() {
            MultipartFile file = new MockMultipartFile("file", "binary.csv", "text/csv", new byte[]{0, 1, 2, 0, 0, 3, 0, 0, 0});

            assertEquals("Upload file content is not text CSV", refusal(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(file)));
        }

        @Test
        void aWorkbookCalledCsvIsNotText() throws Exception {
            MultipartFile file = new MockMultipartFile("file", "really.csv", "text/csv", xlsxLikeZip());

            assertEquals("Upload file content is not text CSV", refusal(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(file)));
        }

        @Test
        void textThatIsNotUtf8IsRefused() {
            byte[] latin1 = "name,city\nJosé,Zürich\n".getBytes(StandardCharsets.ISO_8859_1);
            MultipartFile file = new MockMultipartFile("file", "latin1.csv", "text/csv", latin1);

            assertEquals("Upload file content is not UTF-8 CSV", refusal(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(file)));
        }

        @Test
        void aCharacterCutByTheEndOfTheInspectedPrefixDoesNotMakeAValidFileInvalid() {
            // 4,096 bytes are inspected: put a 3-byte character across that boundary
            StringBuilder text = new StringBuilder("name\n");
            while (text.toString().getBytes(StandardCharsets.UTF_8).length < 4094) text.append('a');
            text.append("한").append("\n".repeat(50)); // the character starts at byte 4094 and ends at 4096
            byte[] bytes = text.toString().getBytes(StandardCharsets.UTF_8);
            assertTrue(bytes.length > 4096);
            MultipartFile file = new MockMultipartFile("file", "long.csv", "text/csv", bytes);

            assertSame(file, ExcelKitMultipartFile.requireExcelOrCsvContent(file));
        }

        @Test
        void aStreamThatFailsWhileItIsInspectedIsReportedAsSuch() {
            ExcelKitUploadException e = assertThrows(ExcelKitUploadException.class,
                    () -> ExcelKitMultipartFile.requireExcelOrCsvContent(StubMultipartFile.unreadable("data.csv")));

            assertEquals("Failed to inspect upload file: data.csv", e.getMessage());
            assertEquals("read failed", e.getCause().getMessage());
        }

        @Test
        void aFileThatClaimsToHaveContentButHasNoneIsEmpty() {
            // its size says 10 bytes, its stream delivers none
            var csv = new StubMultipartFile("lying.csv", 10, () -> new java.io.ByteArrayInputStream(new byte[0]));
            var xlsx = new StubMultipartFile("lying.xlsx", 10, () -> new java.io.ByteArrayInputStream(new byte[0]));

            assertEquals("Upload file is empty", refusal(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(csv)));
            assertEquals("Upload file content is not a valid XLSX/XLSM package",
                    refusal(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(xlsx)));
        }

        @Test
        void anExtensionOtherThanExcelOrCsvIsRefusedBeforeTheContentIsRead() {
            MultipartFile file = new MockMultipartFile("file", "notes.txt", "text/plain", "a".getBytes(StandardCharsets.UTF_8));

            assertEquals("Unsupported upload file extension. Allowed extensions: csv, xlsm, xlsx",
                    refusal(() -> ExcelKitMultipartFile.requireExcelOrCsvContent(file)));
        }
    }
}
