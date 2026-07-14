package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.core.ExcelKitSchema;
import io.github.dornol.excelkit.core.ReadLimits;
import io.github.dornol.excelkit.core.ReadLimitExceededException;
import io.github.dornol.excelkit.core.TabularFileType;
import org.junit.jupiter.api.Test;
import org.springframework.mock.web.MockMultipartFile;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelKitUploadTest {

    @Test
    void csv_opensMultipartFileAndCollectsUploadResult() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name,Price\nNotebook,1200\nPen,bad\n".getBytes(StandardCharsets.UTF_8));

        UploadResult<Product> result = ExcelKitUpload.csv(file, (inputStream, consumer) -> CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .column("Price", (p, cell) -> p.price = cell.asInt())
                .read(inputStream, consumer));

        assertEquals("CSV", result.type());
        assertEquals(1, result.successCount());
        assertEquals(1, result.errorCount());
        assertEquals("Notebook", result.rows().getFirst().name);
    }

    @Test
    void read_usesCustomTypeLabel() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name\nNotebook\n".getBytes(StandardCharsets.UTF_8));

        UploadResult<Product> result = ExcelKitUpload.read("Products", file, (inputStream, consumer) -> CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .read(inputStream, consumer));

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
    void read_closesMultipartStreamExactlyOnce() {
        MultipartFile file = new MockMultipartFile(
                "file", "..\\bad/\r\nname.csv", "text/csv", new byte[]{1}) {
            @Override
            public InputStream getInputStream() {
                return new FailOnSecondCloseInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8));
            }
        };

        UploadResult<Product> result = ExcelKitUpload.csv(file, (inputStream, consumer) -> CsvReader.setter(Product::new)
                .column("Name", (p, cell) -> p.name = cell.asString())
                .read(inputStream, consumer));

        assertEquals(1, result.successCount());
    }

    @Test
    void detectedUploadRejectsContentMismatch() {
        MockMultipartFile file = new MockMultipartFile("file", "fake.xlsx", "application/octet-stream",
                "Name\nA\n".getBytes(StandardCharsets.UTF_8));
        ExcelKitUploadException error = assertThrows(ExcelKitUploadException.class, () ->
                ExcelKitUpload.readDetected("Excel", TabularFileType.XLSX, file, (input, consumer) -> {}));
        assertFalse(error.getMessage().isBlank());
    }

    @Test
    void schemaUploadAppliesCoreReadLimits() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name\nNotebook\n".getBytes(StandardCharsets.UTF_8));
        ExcelKitSchema<Product> schema = ExcelKitSchema.<Product>builder()
                .column("Name", p -> p.name, (p, cell) -> p.name = cell.asString()).build();
        assertThrows(ReadLimitExceededException.class, () -> ExcelKitUpload.csv(file, schema,
                Product::new, null, new ReadLimits(3, -1, -1, -1)));
    }

    @Test
    void detectedUploadClosesUnderlyingStreamExactlyOnce() {
        MultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv", new byte[]{1}) {
            @Override public InputStream getInputStream() {
                return new FailOnSecondCloseInputStream("Name\nAlice\n".getBytes(StandardCharsets.UTF_8));
            }
        };
        UploadResult<Product> result = ExcelKitUpload.readDetected("CSV", TabularFileType.CSV, file,
                (input, consumer) -> CsvReader.setter(Product::new)
                        .column("Name", (p, cell) -> p.name = cell.asString()).read(input, consumer));
        assertEquals(1, result.successCount());
    }

    @Test
    void uploadCollectionLimitsPreserveCountsAndMarkTruncation() {
        MockMultipartFile file = new MockMultipartFile("file", "products.csv", "text/csv",
                "Name,Price\nA,1\nB,2\nC,bad\nD,bad\n".getBytes(StandardCharsets.UTF_8));
        UploadResult<Product> result = ExcelKitUpload.read("CSV", file, (input, consumer) ->
                CsvReader.setter(Product::new)
                        .column("Name", (p, cell) -> p.name = cell.asString())
                        .column("Price", (p, cell) -> p.price = cell.asInt()).read(input, consumer),
                new UploadCollectionLimits(1, 1));
        assertEquals(2, result.successCount());
        assertEquals(2, result.errorCount());
        assertEquals(1, result.rows().size());
        assertEquals(1, result.errors().size());
        assertTrue(result.rowsTruncated());
        assertTrue(result.errorsTruncated());
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
