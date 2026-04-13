package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ExcelColor;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.csv.CsvHandler;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link FileHandler} — the common interface introduced in v0.11.0 so that
 * {@link ExcelHandler} and {@link CsvHandler} can be consumed polymorphically (for
 * example, from a single Spring controller method that returns either format).
 *
 * <p>These tests live in the {@code shared} package (alongside the interface) rather
 * than in one of the format-specific packages, to make clear that the contract is
 * cross-format.
 */
class FileHandlerTest {

    @Test
    @DisplayName("ExcelHandler is assignable to FileHandler")
    void excelHandler_isFileHandler() {
        FileHandler handler = ExcelWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("x"));
        assertNotNull(handler);
        assertInstanceOf(ExcelHandler.class, handler);
    }

    @Test
    @DisplayName("CsvHandler is assignable to FileHandler")
    void csvHandler_isFileHandler() {
        FileHandler handler = CsvWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("x"));
        assertNotNull(handler);
        assertInstanceOf(CsvHandler.class, handler);
    }

    @Test
    @DisplayName("FileHandler.write produces valid Excel bytes when the concrete type is ExcelHandler")
    void write_viaInterface_producesExcelBytes() throws IOException {
        FileHandler handler = ExcelWriter.<String>create().headerColor(ExcelColor.STEEL_BLUE)
                .column("Name", s -> s)
                .write(Stream.of("Alice", "Bob"));

        var out = new ByteArrayOutputStream();
        handler.writeTo(out);

        byte[] bytes = out.toByteArray();
        assertTrue(bytes.length > 0);
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(bytes))) {
            assertEquals("Name", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertEquals("Bob", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    @DisplayName("FileHandler.write produces valid CSV bytes when the concrete type is CsvHandler")
    void write_viaInterface_producesCsvBytes() throws IOException {
        FileHandler handler = CsvWriter.<String>create()
                .column("Name", s -> s)
                .write(Stream.of("Alice", "Bob"));

        var out = new ByteArrayOutputStream();
        handler.writeTo(out);

        String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
        String[] lines = csv.split("\r?\n");
        assertEquals("Name", lines[0]);
        assertEquals("Alice", lines[1]);
        assertEquals("Bob", lines[2]);
    }

    @Test
    @DisplayName("A single polymorphic method can handle either format")
    void polymorphicConsumer_writesEitherFormat() throws IOException {
        // Simulate what a Spring controller would do: receive a FileHandler and hand
        // it off to a lambda. This is exactly the use case the interface was introduced for.
        FileHandler excelH = ExcelWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("excel-payload"));
        FileHandler csvH = CsvWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("csv-payload"));

        byte[] excelBytes = writeToBytes(excelH);
        byte[] csvBytes = writeToBytes(csvH);

        assertTrue(excelBytes.length > 0);
        assertTrue(csvBytes.length > 0);
        // ZIP (XLSX) signature
        assertEquals(0x50, excelBytes[0] & 0xFF);
        assertEquals(0x4B, excelBytes[1] & 0xFF);
        // CSV starts with BOM by default
        String csv = new String(csvBytes, StandardCharsets.UTF_8);
        assertTrue(csv.contains("csv-payload"));
    }

    private static byte[] writeToBytes(FileHandler handler) throws IOException {
        try (ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            handler.writeTo(out);
            return out.toByteArray();
        }
    }

    @Test
    @DisplayName("FileHandler interface is accepted as OutputStream target via method reference")
    void methodReference_asConsumer() throws IOException {
        FileHandler handler = ExcelWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("x"));

        // This is the Spring StreamingResponseBody pattern:
        //   .body(handler::writeTo)
        // which requires handler::writeTo to conform to OutputStream -> void (throws IOException).
        // If FileHandler.write ever changes signature, this test fails to compile.
        ThrowingOutputStreamConsumer sink = handler::writeTo;
        var out = new ByteArrayOutputStream();
        sink.accept(out);
        assertTrue(out.size() > 0);
    }

    @FunctionalInterface
    private interface ThrowingOutputStreamConsumer {
        void accept(OutputStream out) throws IOException;
    }

    @Test
    @DisplayName("ExcelHandler and CsvHandler are final (closed hierarchy)")
    void handlersAreFinal() {
        assertTrue(java.lang.reflect.Modifier.isFinal(ExcelHandler.class.getModifiers()),
                "ExcelHandler must be final so the FileHandler hierarchy stays closed");
        assertTrue(java.lang.reflect.Modifier.isFinal(CsvHandler.class.getModifiers()),
                "CsvHandler must be final so the FileHandler hierarchy stays closed");
    }
}
