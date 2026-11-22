package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.CellData;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.*;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class ExcelReadProgressTest {

    // ========================================================================
    // Excel read progress callback
    // ========================================================================
    @Test
    void readProgress_shouldFireAtCorrectIntervals() throws IOException {
        Path file = tempDir.resolve("progress.xlsx");
        createExcelFileWithRows(file, 10);

        List<Long> progressCounts = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(SimplePerson::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .column((p, cell) -> p.age = cell.asInt())
                    .onProgress(3, (count, cursor) -> progressCounts.add(count))
                    .read(is, r -> {});
        }

        assertEquals(List.of(3L, 6L, 9L), progressCounts);
    }

    @Test
    void readProgress_shouldNotFireWhenIntervalNotReached() throws IOException {
        Path file = tempDir.resolve("progress-small.xlsx");
        createExcelFileWithRows(file, 2);

        List<Long> progressCounts = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(SimplePerson::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .column((p, cell) -> p.age = cell.asInt())
                    .onProgress(100, (count, cursor) -> progressCounts.add(count))
                    .read(is, r -> {});
        }

        assertTrue(progressCounts.isEmpty());
    }

    @Test
    void readProgress_invalidInterval_shouldThrow() {
        assertThrows(IllegalArgumentException.class, () ->
                new ExcelReader<>(SimplePerson::new, null)
                        .onProgress(0, (c, cur) -> {}));
    }

    @Test
    void readProgress_viaBuilderChain() throws IOException {
        Path file = tempDir.resolve("progress-chain.xlsx");
        createExcelFileWithRows(file, 6);

        List<Long> counts = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(SimplePerson::new, null)
                    .column((SimplePerson p, CellData cell) -> p.name = cell.asString())
                    .column((SimplePerson p, CellData cell) -> p.age = cell.asInt())
                    .onProgress(2, (count, cursor) -> counts.add(count))
                    .read(is, r -> {});
        }

        assertEquals(List.of(2L, 4L, 6L), counts);
    }

    @TempDir
    Path tempDir;

    // ========================================================================
    // Helpers
    // ========================================================================
    private void createExcelFileWithRows(Path filePath, int rowCount) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            Row header = sheet.createRow(0);
            header.createCell(0).setCellValue("Name");
            header.createCell(1).setCellValue("Age");
            for (int i = 0; i < rowCount; i++) {
                Row row = sheet.createRow(i + 1);
                row.createCell(0).setCellValue("Person" + i);
                row.createCell(1).setCellValue(20 + i);
            }
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    public static class SimplePerson {
        String name;
        Integer age;
    }
}
