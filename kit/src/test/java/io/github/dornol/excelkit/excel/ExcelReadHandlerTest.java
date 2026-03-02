package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.shared.CellData;
import io.github.dornol.excelkit.shared.ReadResult;
import jakarta.validation.Validation;
import jakarta.validation.Validator;
import jakarta.validation.constraints.Max;
import jakarta.validation.constraints.Min;
import jakarta.validation.constraints.NotBlank;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.hibernate.validator.messageinterpolation.ParameterMessageInterpolator;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.function.BiConsumer;
import java.util.function.Consumer;
import java.util.function.Supplier;

import static org.junit.jupiter.api.Assertions.*;
import static org.junit.jupiter.api.Assertions.assertFalse;

/**
 * Tests for {@link ExcelReadHandler} class.
 */
class ExcelReadHandlerTest {

    @TempDir
    Path tempDir;
    
    private Path excelFile;
    private Validator validator;
    
    @BeforeEach
    void setUp() throws IOException {
        // Create a validator with ParameterMessageInterpolator to avoid EL dependency
        validator = Validation.byDefaultProvider()
                .configure()
                .messageInterpolator(new ParameterMessageInterpolator())
                .buildValidatorFactory()
                .getValidator();
        
        // Create a test Excel file
        excelFile = tempDir.resolve("test.xlsx");
        createTestExcelFile(excelFile);
    }
    
    @Test
    void read_shouldReadExcelFileCorrectly() throws IOException {
        // Arrange
        List<ExcelReadColumn<TestPerson>> columns = new ArrayList<>();
        columns.add(new ExcelReadColumn<>(createNameSetter()));
        columns.add(new ExcelReadColumn<>(createAgeSetter()));
        
        Supplier<TestPerson> instanceSupplier = TestPerson::new;
        
        List<TestPerson> results = new ArrayList<>();
        Consumer<ReadResult<TestPerson>> consumer = result -> {
            if (result.success()) {
                results.add(result.data());
            }
        };
        
        // Act
        try (InputStream is = Files.newInputStream(excelFile)) {
            ExcelReadHandler<TestPerson> handler = new ExcelReadHandler<>(is, columns, instanceSupplier, validator);
            handler.read(consumer);
        }
        
        // Assert
        assertEquals(3, results.size(), "Should read 3 valid records");
        
        TestPerson person1 = results.get(0);
        assertEquals("Alice", person1.getName(), "First person name should be Alice");
        assertEquals(30, person1.getAge(), "First person age should be 30");
        
        TestPerson person2 = results.get(1);
        assertEquals("Bob", person2.getName(), "Second person name should be Bob");
        assertEquals(25, person2.getAge(), "Second person age should be 25");
        
        TestPerson person3 = results.get(2);
        assertEquals("Charlie", person3.getName(), "Third person name should be Charlie");
        assertEquals(35, person3.getAge(), "Third person age should be 35");
    }
    
    @Test
    void read_shouldValidateData() throws IOException {
        // Arrange
        List<ExcelReadColumn<TestPerson>> columns = new ArrayList<>();
        columns.add(new ExcelReadColumn<>(createNameSetter()));
        columns.add(new ExcelReadColumn<>(createAgeSetter()));
        
        Supplier<TestPerson> instanceSupplier = TestPerson::new;
        
        List<TestPerson> validResults = new ArrayList<>();
        List<ReadResult<TestPerson>> invalidResults = new ArrayList<>();
        
        Consumer<ReadResult<TestPerson>> consumer = result -> {
            if (result.success()) {
                validResults.add(result.data());
            } else {
                invalidResults.add(result);
            }
        };
        
        // Create a test file with invalid data
        Path invalidExcelFile = tempDir.resolve("invalid.xlsx");
        createInvalidTestExcelFile(invalidExcelFile);
        
        // Act
        try (InputStream is = Files.newInputStream(invalidExcelFile)) {
            ExcelReadHandler<TestPerson> handler = new ExcelReadHandler<>(is, columns, instanceSupplier, validator);
            handler.read(consumer);
        }
        
        // Assert
        assertEquals(1, validResults.size(), "Should have 1 valid record");
        assertEquals(2, invalidResults.size(), "Should have 2 invalid records");
        
        // Check that we have validation errors
        ReadResult<TestPerson> invalidResult1 = invalidResults.get(0);
        assertFalse(invalidResult1.success(), "First invalid result should have success=false");
        assertFalse(invalidResult1.messages().isEmpty(), "First invalid result should have error messages");
        
        ReadResult<TestPerson> invalidResult2 = invalidResults.get(1);
        assertFalse(invalidResult2.success(), "Second invalid result should have success=false");
        assertFalse(invalidResult2.messages().isEmpty(), "Second invalid result should have error messages");
    }
    
    @Test
    void read_shouldReadSecondSheet() throws IOException {
        Path multiSheetFile = tempDir.resolve("multi-sheet.xlsx");
        createMultiSheetExcelFile(multiSheetFile);

        List<TestPerson> results = new ArrayList<>();
        Consumer<ReadResult<TestPerson>> consumer = result -> {
            if (result.success()) {
                results.add(result.data());
            }
        };

        try (InputStream is = Files.newInputStream(multiSheetFile)) {
            new ExcelReader<>(TestPerson::new, validator)
                    .sheetIndex(1)
                    .column(createNameSetter())
                    .column(createAgeSetter())
                    .build(is)
                    .read(consumer);
        }

        assertEquals(2, results.size());
        assertEquals("Dave", results.get(0).getName());
        assertEquals(40, results.get(0).getAge());
        assertEquals("Eve", results.get(1).getName());
        assertEquals(28, results.get(1).getAge());
    }

    @Test
    void read_shouldReadSecondSheetViaBuilderOverload() throws IOException {
        Path multiSheetFile = tempDir.resolve("multi-sheet2.xlsx");
        createMultiSheetExcelFile(multiSheetFile);

        List<TestPerson> results = new ArrayList<>();
        Consumer<ReadResult<TestPerson>> consumer = result -> {
            if (result.success()) {
                results.add(result.data());
            }
        };

        try (InputStream is = Files.newInputStream(multiSheetFile)) {
            new ExcelReader<>(TestPerson::new, validator)
                    .column(createNameSetter())
                    .column(createAgeSetter())
                    .build(is, 1)
                    .read(consumer);
        }

        assertEquals(2, results.size());
        assertEquals("Dave", results.get(0).getName());
    }

    @Test
    void read_shouldThrowForInvalidSheetIndex() throws IOException {
        Path singleSheetFile = tempDir.resolve("single-sheet.xlsx");
        createTestExcelFile(singleSheetFile);

        try (InputStream is = Files.newInputStream(singleSheetFile)) {
            ExcelReadHandler<TestPerson> handler = new ExcelReader<>(TestPerson::new, validator)
                    .sheetIndex(5)
                    .column(createNameSetter())
                    .column(createAgeSetter())
                    .build(is);

            assertThrows(ExcelReadException.class, () -> handler.read(r -> {}));
        }
    }

    @Test
    void constructor_shouldThrowForNegativeSheetIndex() {
        assertThrows(IllegalArgumentException.class, () -> {
            List<ExcelReadColumn<TestPerson>> columns = new ArrayList<>();
            columns.add(new ExcelReadColumn<>(createNameSetter()));
            new ExcelReadHandler<>(InputStream.nullInputStream(), columns, TestPerson::new, null, -1);
        });
    }

    @Test
    void read_shouldSkipRowsBeforeHeaderRowIndex() throws IOException {
        Path file = tempDir.resolve("header-offset.xlsx");
        createExcelFileWithHeaderOffset(file);

        List<TestPerson> results = new ArrayList<>();
        Consumer<ReadResult<TestPerson>> consumer = result -> {
            if (result.success()) {
                results.add(result.data());
            }
        };

        try (InputStream is = Files.newInputStream(file)) {
            new ExcelReader<>(TestPerson::new, validator)
                    .headerRowIndex(2)
                    .column(createNameSetter())
                    .column(createAgeSetter())
                    .build(is)
                    .read(consumer);
        }

        assertEquals(2, results.size());
        assertEquals("Alice", results.get(0).getName());
        assertEquals(30, results.get(0).getAge());
        assertEquals("Bob", results.get(1).getName());
        assertEquals(25, results.get(1).getAge());
    }

    @Test
    void constructor_shouldThrowForSheetIndexExceeding255() {
        assertThrows(IllegalArgumentException.class, () -> {
            List<ExcelReadColumn<TestPerson>> columns = new ArrayList<>();
            columns.add(new ExcelReadColumn<>(createNameSetter()));
            new ExcelReadHandler<>(InputStream.nullInputStream(), columns, TestPerson::new, null, 256);
        }, "sheetIndex > 255 should throw IllegalArgumentException");
    }

    @Test
    void constructor_shouldThrowForNegativeHeaderRowIndex() {
        assertThrows(IllegalArgumentException.class, () -> {
            List<ExcelReadColumn<TestPerson>> columns = new ArrayList<>();
            columns.add(new ExcelReadColumn<>(createNameSetter()));
            new ExcelReadHandler<>(InputStream.nullInputStream(), columns, TestPerson::new, null, 0, -1);
        });
    }

    /**
     * Creates a setter function for the name field.
     */
    private BiConsumer<TestPerson, CellData> createNameSetter() {
        return (person, cellData) -> person.setName(cellData.asString());
    }
    
    /**
     * Creates a setter function for the age field.
     */
    private BiConsumer<TestPerson, CellData> createAgeSetter() {
        return (person, cellData) -> person.setAge(cellData.asInt());
    }
    
    /**
     * Creates a test Excel file with valid data.
     */
    private void createTestExcelFile(Path filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            
            // Header row
            Row headerRow = sheet.createRow(0);
            Cell nameHeaderCell = headerRow.createCell(0);
            nameHeaderCell.setCellValue("Name");
            Cell ageHeaderCell = headerRow.createCell(1);
            ageHeaderCell.setCellValue("Age");
            
            // Data rows
            createDataRow(sheet, 1, "Alice", 30);
            createDataRow(sheet, 2, "Bob", 25);
            createDataRow(sheet, 3, "Charlie", 35);
            
            // Write to file
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }
    
    /**
     * Creates a test Excel file with some invalid data.
     */
    private void createInvalidTestExcelFile(Path filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");
            
            // Header row
            Row headerRow = sheet.createRow(0);
            Cell nameHeaderCell = headerRow.createCell(0);
            nameHeaderCell.setCellValue("Name");
            Cell ageHeaderCell = headerRow.createCell(1);
            ageHeaderCell.setCellValue("Age");
            
            // Data rows - one valid, two invalid
            createDataRow(sheet, 1, "Valid", 30); // Valid
            createDataRow(sheet, 2, "", 25);      // Invalid: blank name
            createDataRow(sheet, 3, "TooOld", 150); // Invalid: age > 100
            
            // Write to file
            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }
    
    /**
     * Creates a test Excel file with multiple sheets.
     */
    private void createMultiSheetExcelFile(Path filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            // Sheet 1
            Sheet sheet1 = workbook.createSheet("Sheet1");
            Row header1 = sheet1.createRow(0);
            header1.createCell(0).setCellValue("Name");
            header1.createCell(1).setCellValue("Age");
            createDataRow(sheet1, 1, "Alice", 30);
            createDataRow(sheet1, 2, "Bob", 25);
            createDataRow(sheet1, 3, "Charlie", 35);

            // Sheet 2
            Sheet sheet2 = workbook.createSheet("Sheet2");
            Row header2 = sheet2.createRow(0);
            header2.createCell(0).setCellValue("Name");
            header2.createCell(1).setCellValue("Age");
            createDataRow(sheet2, 1, "Dave", 40);
            createDataRow(sheet2, 2, "Eve", 28);

            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    /**
     * Creates a test Excel file with metadata rows before the header.
     * Row 0: "Report Title"
     * Row 1: "Generated: 2025-01-01"
     * Row 2: Header (Name, Age)
     * Row 3+: Data
     */
    private void createExcelFileWithHeaderOffset(Path filePath) throws IOException {
        try (Workbook workbook = new XSSFWorkbook()) {
            Sheet sheet = workbook.createSheet("Test");

            // Metadata rows
            Row metaRow0 = sheet.createRow(0);
            metaRow0.createCell(0).setCellValue("Report Title");
            Row metaRow1 = sheet.createRow(1);
            metaRow1.createCell(0).setCellValue("Generated: 2025-01-01");

            // Header row at index 2
            Row headerRow = sheet.createRow(2);
            headerRow.createCell(0).setCellValue("Name");
            headerRow.createCell(1).setCellValue("Age");

            // Data rows
            createDataRow(sheet, 3, "Alice", 30);
            createDataRow(sheet, 4, "Bob", 25);

            try (FileOutputStream fos = new FileOutputStream(filePath.toFile())) {
                workbook.write(fos);
            }
        }
    }

    private void createDataRow(Sheet sheet, int rowNum, String name, int age) {
        Row row = sheet.createRow(rowNum);
        Cell nameCell = row.createCell(0);
        nameCell.setCellValue(name);
        Cell ageCell = row.createCell(1);
        ageCell.setCellValue(age);
    }
    
    /**
     * Test class for Excel reading tests.
     */
    public static class TestPerson {
        @NotBlank
        private String name;
        
        @Min(1)
        @Max(100)
        private int age;
        
        public String getName() {
            return name;
        }
        
        public void setName(String name) {
            this.name = name;
        }
        
        public int getAge() {
            return age;
        }
        
        public void setAge(int age) {
            this.age = age;
        }
    }
}