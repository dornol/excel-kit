package io.github.dornol.excelkit.excel;

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

import static org.junit.jupiter.api.Assertions.assertEquals;
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
        Consumer<ExcelReadResult<TestPerson>> consumer = result -> {
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
        List<ExcelReadResult<TestPerson>> invalidResults = new ArrayList<>();
        
        Consumer<ExcelReadResult<TestPerson>> consumer = result -> {
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
        ExcelReadResult<TestPerson> invalidResult1 = invalidResults.get(0);
        assertFalse(invalidResult1.success(), "First invalid result should have success=false");
        assertFalse(invalidResult1.messages().isEmpty(), "First invalid result should have error messages");
        
        ExcelReadResult<TestPerson> invalidResult2 = invalidResults.get(1);
        assertFalse(invalidResult2.success(), "Second invalid result should have success=false");
        assertFalse(invalidResult2.messages().isEmpty(), "Second invalid result should have error messages");
    }
    
    /**
     * Creates a setter function for the name field.
     */
    private BiConsumer<TestPerson, ExcelCellData> createNameSetter() {
        return (person, cellData) -> person.setName(cellData.asString());
    }
    
    /**
     * Creates a setter function for the age field.
     */
    private BiConsumer<TestPerson, ExcelCellData> createAgeSetter() {
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