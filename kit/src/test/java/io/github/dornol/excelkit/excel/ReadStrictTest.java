package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.read.ExcelReader;
import jakarta.validation.constraints.NotBlank;
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

class ReadStrictTest {

    // ========================================================================
    // readStrict - row number in error message
    // ========================================================================

    @Test
    void readStrict_successShouldNotThrow() throws IOException {
        Path file = tempDir.resolve("valid.xlsx");
        try (Workbook wb = new XSSFWorkbook()) {
            Sheet sheet = wb.createSheet("Test");
            sheet.createRow(0).createCell(0).setCellValue("Name");
            sheet.createRow(1).createCell(0).setCellValue("Alice");
            sheet.createRow(2).createCell(0).setCellValue("Bob");
            try (FileOutputStream fos = new FileOutputStream(file.toFile())) {
                wb.write(fos);
            }
        }

        List<String> names = new ArrayList<>();
        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(ValidatedPerson::new, null)
                    .column((p, cell) -> p.name = cell.asString())
                    .readStrict(is, p -> names.add(p.name));
        }

        assertEquals(List.of("Alice", "Bob"), names);
    }

    @TempDir
    Path tempDir;

    public static class ValidatedPerson {
        @NotBlank(message = "Name must not be blank")
        String name;
    }
}
