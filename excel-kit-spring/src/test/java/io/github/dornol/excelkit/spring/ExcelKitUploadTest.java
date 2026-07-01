package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.core.ExcelKitSchema;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;

class ExcelKitUploadTest {

    @Test
    void csv_opensMultipartFileAndCollectsUploadResult() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name,Price\nNotebook,1200\nPen,bad\n".getBytes(StandardCharsets.UTF_8));

        UploadResult<Product> result = ExcelKitUpload.csv(file, inputStream -> CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Price", (p, cell) -> p.price = cell.asInt())
                .build(inputStream));

        assertEquals("CSV", result.type());
        assertEquals(1, result.successCount());
        assertEquals(1, result.errorCount());
        assertEquals("Notebook", result.rows().getFirst().name);
    }

    @Test
    void read_usesCustomTypeLabel() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name\nNotebook\n".getBytes(StandardCharsets.UTF_8));

        UploadResult<Product> result = ExcelKitUpload.read("Products", file, inputStream -> CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .build(inputStream));

        assertEquals("Products", result.type());
        assertEquals(1, result.successCount());
        assertEquals(0, result.errorCount());
    }

    @Test
    void csv_schemaShortcut_opensMultipartFileAndCollectsUploadResult() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name,Price\nNotebook,1200\n".getBytes(StandardCharsets.UTF_8));
        ExcelKitSchema<Product> schema = ExcelKitSchema.<Product>builder()
                .column("Name", p -> p.name, (p, cell) -> p.name = cell.asString())
                .column("Price", p -> p.price, (p, cell) -> p.price = cell.asInt())
                .build();

        UploadResult<Product> result = ExcelKitUpload.csv(file, schema, Product::new);

        assertEquals("CSV", result.type());
        assertEquals(1, result.successCount());
        assertEquals("Notebook", result.rows().getFirst().name);
    }

    @Test
    void read_sanitizesFilenameWhenUploadStreamCloseFails() {
        MultipartFile file = new MockMultipartFile(
                "file", "..\\bad/\r\nname.csv", "text/csv", new byte[]{1}) {
            @Override
            public InputStream getInputStream() {
                return new FailOnSecondCloseInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8));
            }
        };

        ExcelKitUploadException exception = assertThrows(ExcelKitUploadException.class,
                () -> ExcelKitUpload.csv(file, inputStream -> CsvReader.setter(Product::new)
                        .column("Name", (p, cell) -> p.name = cell.asString())
                        .build(inputStream)));

        assertFalse(exception.getMessage().contains("\r"));
        assertFalse(exception.getMessage().contains("\n"));
        assertFalse(exception.getMessage().contains("\\"));
        assertFalse(exception.getMessage().contains("/"));
    }

    static class Product {
        String name;
        Integer price;
    }

    private static final class FailOnSecondCloseInputStream extends InputStream {
        private final byte[] data;
        private int index;
        private int closeCount;

        private FailOnSecondCloseInputStream(byte[] data) {
            this.data = data;
        }

        @Override
        public int read() {
            if (index >= data.length) {
                return -1;
            }
            return data[index++] & 0xFF;
        }

        @Override
        public void close() throws IOException {
            closeCount++;
            if (closeCount > 1) {
                throw new IOException("close failed");
            }
        }
    }
}
