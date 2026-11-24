package io.github.dornol.excelkit.spring;

import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertSame;
import static org.junit.jupiter.api.Assertions.assertThrows;

class ExcelKitMultipartFileTest {

    @Test
    void open_returnsMultipartFileInputStream() throws Exception {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample.csv", "text/csv", "a,b\n".getBytes(StandardCharsets.UTF_8));

        try (var inputStream = ExcelKitMultipartFile.open(file)) {
            assertEquals("a,b\n", new String(inputStream.readAllBytes(), StandardCharsets.UTF_8));
        }
    }

    @Test
    void open_rejectsEmptyFile() {
        MockMultipartFile file = new MockMultipartFile("file", "empty.csv", "text/csv", new byte[0]);

        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.open(file));
    }

    @Test
    void requireSizeAtMost_returnsFileWhenWithinLimit() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample.csv", "text/csv", "a,b\n".getBytes(StandardCharsets.UTF_8));

        assertSame(file, ExcelKitMultipartFile.requireSizeAtMost(file, 4));
    }

    @Test
    void requireSizeAtMost_rejectsOversizedFile() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample.csv", "text/csv", "a,b\n".getBytes(StandardCharsets.UTF_8));

        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireSizeAtMost(file, 3));
    }

    @Test
    void requireExtension_returnsFileForAllowedExtensionIgnoringCaseAndDot() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample.XLSX", "application/octet-stream", "data".getBytes(StandardCharsets.UTF_8));

        assertSame(file, ExcelKitMultipartFile.requireExtension(file, ".xlsx", "csv"));
    }

    @Test
    void requireExtension_rejectsUnsupportedExtension() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample.txt", "text/plain", "data".getBytes(StandardCharsets.UTF_8));

        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExtension(file, "xlsx", "csv"));
    }

    @Test
    void requireExtension_rejectsMissingFilenameExtension() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample", "application/octet-stream", "data".getBytes(StandardCharsets.UTF_8));

        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExtension(file, "xlsx", "csv"));
    }

    @Test
    void requireExcelOrCsv_acceptsCsvAndExcelExtensions() {
        MockMultipartFile csv = new MockMultipartFile(
                "file", "sample.csv", "text/csv", "data".getBytes(StandardCharsets.UTF_8));
        MockMultipartFile excel = new MockMultipartFile(
                "file", "sample.xlsm", "application/octet-stream", "data".getBytes(StandardCharsets.UTF_8));

        assertSame(csv, ExcelKitMultipartFile.requireExcelOrCsv(csv));
        assertSame(excel, ExcelKitMultipartFile.requireExcelOrCsv(excel));
    }

    @Test
    void requireExcelOrCsvContent_acceptsCsvTextAndXlsxZipHeader() {
        MockMultipartFile csv = new MockMultipartFile(
                "file", "sample.csv", "text/csv", "Name\nAlice\n".getBytes(StandardCharsets.UTF_8));
        MockMultipartFile excel = new MockMultipartFile(
                "file", "sample.xlsx", "application/octet-stream", new byte[]{'P', 'K', 3, 4});

        assertSame(csv, ExcelKitMultipartFile.requireExcelOrCsvContent(csv));
        assertSame(excel, ExcelKitMultipartFile.requireExcelOrCsvContent(excel));
    }

    @Test
    void requireExcelOrCsvContent_rejectsMismatchedXlsxContent() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample.xlsx", "application/octet-stream", "not zip".getBytes(StandardCharsets.UTF_8));

        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(file));
    }

    @Test
    void requireExcelOrCsvContent_rejectsBinaryCsvContent() {
        MockMultipartFile file = new MockMultipartFile(
                "file", "sample.csv", "text/csv", new byte[]{'A', 0, 'B'});

        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(file));
    }

    @Test
    void requireExcelOrCsvContent_acceptsMultibyteCsvWhenThePrefixEndsInsideACharacter() {
        // The 4096-byte inspection prefix lands inside a 3-byte Korean character for some paddings.
        for (int pad = 0; pad < 3; pad++) {
            String csv = "a".repeat(pad) + "이름,값\n" + "가나다,1\n".repeat(2000);
            MockMultipartFile file = new MockMultipartFile(
                    "file", "data.csv", "text/csv", csv.getBytes(StandardCharsets.UTF_8));

            assertSame(file, ExcelKitMultipartFile.requireExcelOrCsvContent(file), "pad=" + pad);
        }
    }

    @Test
    void requireExcelOrCsvContent_acceptsCp949CsvOnlyWhenTheCharsetIsGiven() {
        java.nio.charset.Charset cp949 = java.nio.charset.Charset.forName("MS949");
        for (int pad = 0; pad < 2; pad++) {
            String csv = "a".repeat(pad) + "이름,값\n" + "가나다,1\n".repeat(2000);
            MockMultipartFile file = new MockMultipartFile("file", "data.csv", "text/csv", csv.getBytes(cp949));

            assertSame(file, ExcelKitMultipartFile.requireExcelOrCsvContent(file, cp949), "pad=" + pad);
            assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(file));
        }
    }

    @Test
    void requireExcelOrCsvContent_withCharset_rejectsUtf8BytesAndNull() {
        java.nio.charset.Charset cp949 = java.nio.charset.Charset.forName("MS949");
        MockMultipartFile utf8 = new MockMultipartFile("file", "data.csv", "text/csv",
                "이름,값\n가나다,1\n".getBytes(StandardCharsets.UTF_8));
        MockMultipartFile utf16 = new MockMultipartFile("file", "data.csv", "text/csv",
                "a,b\n1,2\n".getBytes(StandardCharsets.UTF_16));

        // UTF-8 Korean text is not valid MS949 here (a lead byte followed by an invalid trail byte)
        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(utf8, cp949));
        assertSame(utf16, ExcelKitMultipartFile.requireExcelOrCsvContent(utf16, StandardCharsets.UTF_16));
        assertThrows(NullPointerException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(utf8, null));
    }

    @Test
    void requireExcelOrCsvContent_rejectsInvalidUtf8EvenInsideALargeFile() {
        byte[] head = ("Name\n" + "x".repeat(100) + "\n").getBytes(StandardCharsets.UTF_8);
        byte[] bytes = new byte[head.length + 3];
        System.arraycopy(head, 0, bytes, 0, head.length);
        bytes[head.length] = (byte) 0xC3;      // malformed: lead byte followed by an ASCII byte
        bytes[head.length + 1] = 'x';
        bytes[head.length + 2] = '\n';

        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(
                new MockMultipartFile("file", "bad.csv", "text/csv", bytes)));
        // A whole file that ends with an incomplete character is still invalid.
        assertThrows(ExcelKitUploadException.class, () -> ExcelKitMultipartFile.requireExcelOrCsvContent(
                new MockMultipartFile("file", "cut.csv", "text/csv", new byte[]{'a', ',', (byte) 0xEC, (byte) 0x9D})));
    }
}
