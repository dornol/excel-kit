package io.github.dornol.excelkit.spring;

import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;

import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
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
}
