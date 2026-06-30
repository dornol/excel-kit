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
}
