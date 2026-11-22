package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.detect.TabularFileType;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** What the upload helpers do when the stream misbehaves or the content is not what was declared. */
class ExcelKitUploadFailuresTest {

    private static final byte[] CSV = "a\n1\n".getBytes(StandardCharsets.UTF_8);

    private static UploadReader<String> countingReader(List<String> seen) {
        return (in, consumer) -> {
            try {
                seen.add(new String(in.readAllBytes(), StandardCharsets.UTF_8));
            } catch (java.io.IOException e) {
                throw new IllegalStateException(e);
            }
            consumer.accept(ReadResult.success("row"));
        };
    }

    @Test
    void aStreamThatFailsWhenItIsClosedIsAnUploadException() {
        for (String variant : List.of("plain", "limited", "detected")) {
            var file = StubMultipartFile.failingOnClose("data.csv", CSV);
            List<String> seen = new ArrayList<>();

            ExcelKitUploadException e = assertThrows(ExcelKitUploadException.class, () -> {
                switch (variant) {
                    case "plain" -> ExcelKitUpload.read("CSV", file, countingReader(seen));
                    case "limited" -> ExcelKitUpload.read("CSV", file, countingReader(seen), UploadCollectionLimits.UNLIMITED);
                    default -> ExcelKitUpload.readDetected("CSV", TabularFileType.CSV, file, countingReader(seen));
                }
            }, variant);

            assertEquals("Failed to close upload file: data.csv", e.getMessage(), variant);
            assertEquals("close failed", e.getCause().getMessage(), variant);
            assertEquals(1, seen.size(), variant + ": the file was read before it failed to close");
        }
    }

    @Test
    void aFileOfAnotherKindThanExpectedIsRefusedBeforeTheReaderRuns() throws Exception {
        List<String> seen = new ArrayList<>();
        var workbookNamedCsv = new MockMultipartFile("file", "x.csv", "text/csv", workbookBytes());

        ExcelKitUploadException e = assertThrows(ExcelKitUploadException.class,
                () -> ExcelKitUpload.readDetected("CSV", TabularFileType.CSV, workbookNamedCsv, countingReader(seen)));

        assertEquals("Expected CSV content but detected XLSX", e.getMessage());
        assertTrue(seen.isEmpty(), "the reader must not see a file of the wrong kind");
    }

    @Test
    void anEmptyUploadIsRefusedByEveryEntryPoint() {
        var empty = new MockMultipartFile("file", "e.csv", "text/csv", new byte[0]);
        UploadReader<String> reader = countingReader(new ArrayList<>());

        assertEquals("Upload file is empty", assertThrows(ExcelKitUploadException.class, () -> ExcelKitUpload.read("CSV", empty, reader)).getMessage());
        assertEquals("Upload file is empty", assertThrows(ExcelKitUploadException.class,
                () -> ExcelKitUpload.read("CSV", empty, reader, UploadCollectionLimits.UNLIMITED)).getMessage());
        assertEquals("Upload file is empty", assertThrows(ExcelKitUploadException.class, () -> ExcelKitUpload.excel(empty, reader)).getMessage());
        assertEquals("Upload file is empty", assertThrows(ExcelKitUploadException.class,
                () -> ExcelKitUpload.readDetected("CSV", TabularFileType.CSV, empty, reader)).getMessage());
    }

    @Test
    void excelWithAReaderOfYourOwnLabelsTheResultExcel() {
        List<String> seen = new ArrayList<>();
        var file = new MockMultipartFile("file", "book.xlsx", "application/octet-stream", CSV);

        UploadResult<String> result = ExcelKitUpload.excel(file, countingReader(seen));

        assertEquals("Excel", result.type());
        assertEquals(1, result.successCount());
        assertEquals(1, seen.size());
    }

    private static byte[] workbookBytes() throws Exception {
        var out = new java.io.ByteArrayOutputStream();
        io.github.dornol.excelkit.excel.ExcelWriter.<String>create().column("A", s -> s).write(List.of("x")).writeTo(out);
        return out.toByteArray();
    }
}
