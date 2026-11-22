package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.FileHandler;
import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

// ================================================================
// 1. FileHandler.writeTo(Path)
// ================================================================

@DisplayName("FileHandler.writeTo(Path)")
class HandlerToFileTest {

    @Test
    @DisplayName("ExcelHandler.toFile writes a valid .xlsx file")
    void excelHandler_toFile_writesValidXlsx(@TempDir Path tempDir) throws IOException {
        Path target = tempDir.resolve("output.xlsx");

        ExcelWriter.<String>create()
                .column("Name", s -> s)
                .write(Stream.of("Alice", "Bob"))
                .writeTo(target);

        assertTrue(Files.exists(target));
        assertTrue(Files.size(target) > 0);
        try (var wb = new XSSFWorkbook(Files.newInputStream(target))) {
            assertEquals("Name", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
            assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertEquals("Bob", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
        }
    }

    @Test
    @DisplayName("CsvHandler.toFile writes a valid .csv file")
    void csvHandler_toFile_writesValidCsv(@TempDir Path tempDir) throws IOException {
        Path target = tempDir.resolve("output.csv");

        CsvWriter.<String>create()
                .column("Name", s -> s)
                .write(Stream.of("Alice", "Bob"))
                .writeTo(target);

        assertTrue(Files.exists(target));
        String content = Files.readString(target, StandardCharsets.UTF_8).replace("\uFEFF", "");
        String[] lines = content.split("\r?\n");
        assertEquals("Name", lines[0]);
        assertEquals("Alice", lines[1]);
        assertEquals("Bob", lines[2]);
    }

    @Test
    @DisplayName("Calling toFile() twice on ExcelHandler throws (one-shot contract)")
    void excelHandler_toFile_twice_throws(@TempDir Path tempDir) throws IOException {
        FileHandler handler = ExcelWriter.<String>create()
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        handler.writeTo(tempDir.resolve("first.xlsx"));

        assertThrows(ExcelWriteException.class,
                () -> handler.writeTo(tempDir.resolve("second.xlsx")));
    }

    @Test
    @DisplayName("Calling toFile() twice on CsvHandler throws (one-shot contract)")
    void csvHandler_toFile_twice_throws(@TempDir Path tempDir) throws IOException {
        FileHandler handler = CsvWriter.<String>create()
                .column("Name", s -> s)
                .write(Stream.of("Alice"));

        handler.writeTo(tempDir.resolve("first.csv"));

        assertThrows(CsvWriteException.class,
                () -> handler.writeTo(tempDir.resolve("second.csv")));
    }

    @Test
    @DisplayName("toFile content matches write(OutputStream) content for Excel")
    void excelHandler_toFile_matchesWriteOutput(@TempDir Path tempDir) throws IOException {
        // Write via OutputStream
        var out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Val", s -> s)
                .write(Stream.of("X"))
                .writeTo(out);

        // Write via toFile
        Path target = tempDir.resolve("out.xlsx");
        ExcelWriter.<String>create()
                .column("Val", s -> s)
                .write(Stream.of("X"))
                .writeTo(target);

        // Both should produce valid xlsx with identical data
        try (var wb1 = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
             var wb2 = new XSSFWorkbook(Files.newInputStream(target))) {
            assertEquals(
                    wb1.getSheetAt(0).getRow(1).getCell(0).getStringCellValue(),
                    wb2.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }
    }

    @Test
    @DisplayName("toFile content matches write(OutputStream) content for CSV")
    void csvHandler_toFile_matcheWriteOutput(@TempDir Path tempDir) throws IOException {
        // Write via OutputStream
        var out = new ByteArrayOutputStream();
        CsvWriter.<String>create()
                .column("Val", s -> s)
                .write(Stream.of("X"))
                .writeTo(out);

        // Write via toFile
        Path target = tempDir.resolve("out.csv");
        CsvWriter.<String>create()
                .column("Val", s -> s)
                .write(Stream.of("X"))
                .writeTo(target);

        String fromStream = out.toString(StandardCharsets.UTF_8);
        String fromFile = Files.readString(target, StandardCharsets.UTF_8);
        assertEquals(fromStream, fromFile);
    }

    @Test
    @DisplayName("Calling write() after toFile() throws (one-shot)")
    void writeAfterToFile_throws(@TempDir Path tempDir) throws IOException {
        FileHandler handler = ExcelWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("x"));

        handler.writeTo(tempDir.resolve("out.xlsx"));

        assertThrows(ExcelWriteException.class,
                () -> handler.writeTo(new ByteArrayOutputStream()));
    }

    @Test
    @DisplayName("Calling toFile() after write() throws (one-shot)")
    void toFileAfterWrite_throws(@TempDir Path tempDir) throws IOException {
        FileHandler handler = ExcelWriter.<String>create()
                .column("A", s -> s)
                .write(Stream.of("x"));

        handler.writeTo(new ByteArrayOutputStream());

        assertThrows(ExcelWriteException.class,
                () -> handler.writeTo(tempDir.resolve("out.xlsx")));
    }
}
