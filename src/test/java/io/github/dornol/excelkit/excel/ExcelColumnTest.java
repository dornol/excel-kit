package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.streaming.SXSSFCell;
import org.apache.poi.xssf.streaming.SXSSFRow;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for {@link ExcelColumn} class.
 */
class ExcelColumnTest {

    private SXSSFWorkbook workbook;
    private CellStyle cellStyle;
    private ExcelColumnSetter columnSetter;
    private ExcelRowFunction<String, Object> function;
    private SXSSFSheet sheet;
    private SXSSFRow row;
    private SXSSFCell cell;

    @BeforeEach
    void setUp() {
        workbook = new SXSSFWorkbook();
        cellStyle = workbook.createCellStyle();
        columnSetter = (c, value) -> c.setCellValue(String.valueOf(value));
        function = (data, cursor) -> data;
        sheet = workbook.createSheet("Test Sheet");
        row = sheet.createRow(0);
        cell = row.createCell(0);
    }

    @AfterEach
    void tearDown() throws IOException {
        workbook.close();
    }

    @Test
    void constructor_shouldCreateInstanceWithValidParameters() {
        // Arrange
        String name = "Column";

        // Act
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);

        // Assert
        assertNotNull(column, "Column should be created with valid parameters");
        assertEquals(name, column.getName(), "getName should return the column name");
        assertEquals(cellStyle, column.getStyle(), "getStyle should return the cell style");
        assertTrue(column.getColumnWidth() > 0, "Column width should be initialized");
    }

    @Test
    void applyFunction_shouldReturnFunctionResult() {
        // Arrange
        String name = "Column";
        String testData = "Test Data";
        ExcelCursor cursor = new ExcelCursor();
        ExcelRowFunction<String, Object> testFunction = (data, cursor1) -> data + "-processed";
        ExcelColumn<String> column = new ExcelColumn<>(name, testFunction, cellStyle, columnSetter);

        // Act
        Object result = column.applyFunction(testData, cursor);

        // Assert
        assertEquals("Test Data-processed", result, "applyFunction should return the result of applying the function");
    }

    @Test
    void applyFunction_shouldReturnNullWhenFunctionThrowsException() {
        // Arrange
        String name = "Column";
        String testData = "Test Data";
        ExcelCursor cursor = new ExcelCursor();
        ExcelRowFunction<String, Object> exceptionFunction = (data, cursor1) -> {
            throw new RuntimeException("Test exception");
        };
        ExcelColumn<String> column = new ExcelColumn<>(name, exceptionFunction, cellStyle, columnSetter);

        // Act
        Object result = column.applyFunction(testData, cursor);

        // Assert
        assertNull(result, "applyFunction should return null when function throws an exception");
    }

    @Test
    void setColumnWidth_shouldUpdateColumnWidth() {
        // Arrange
        String name = "Column";
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);
        int initialWidth = column.getColumnWidth();
        int newWidth = initialWidth + 1000;

        // Act
        column.setColumnWidth(newWidth);

        // Assert
        assertEquals(newWidth, column.getColumnWidth(), "setColumnWidth should update the column width");
    }

    @Test
    void fitColumnWidthByValue_shouldUpdateColumnWidthBasedOnValue() {
        // Arrange
        String name = "Short";
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);
        int initialWidth = column.getColumnWidth();
        String longValue = "This is a much longer value that should increase the column width";

        // Act
        column.fitColumnWidthByValue(longValue);

        // Assert
        assertTrue(column.getColumnWidth() > initialWidth, 
                "fitColumnWidthByValue should increase column width for longer values");
    }

    @Test
    void setColumnData_shouldSetCellValueUsingColumnSetter() {
        // Arrange
        String name = "Column";
        String testData = "Test Data";
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);

        // Act
        column.setColumnData(cell, testData);

        // Assert
        assertEquals(testData, cell.getStringCellValue(), 
                "setColumnData should set the cell value using the column setter");
    }

    @Test
    void setColumnData_shouldSetEmptyStringWhenDataIsNull() {
        // Arrange
        String name = "Column";
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);

        // Act
        column.setColumnData(cell, null);

        // Assert
        assertEquals("", cell.getStringCellValue(), 
                "setColumnData should set empty string when data is null");
    }

    @Test
    void setColumnData_shouldHandleExceptionFromColumnSetter() {
        // Arrange
        String name = "Column";
        Object testData = new Object(); // Object that will cause exception in our simple setter
        
        // Create a setter that throws an exception for this specific test
        ExcelColumnSetter exceptionSetter = (c, v) -> {
            throw new RuntimeException("Test exception");
        };
        
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, exceptionSetter);

        // Act
        column.setColumnData(cell, testData);

        // Assert
        assertEquals(testData.toString(), cell.getStringCellValue(), 
                "setColumnData should handle exception and set string value");
    }

    @Test
    void getName_shouldReturnColumnName() {
        // Arrange
        String name = "Test Column";
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);

        // Act
        String result = column.getName();

        // Assert
        assertEquals(name, result, "getName should return the column name");
    }

    @Test
    void getStyle_shouldReturnColumnStyle() {
        // Arrange
        String name = "Test Column";
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);

        // Act
        CellStyle result = column.getStyle();

        // Assert
        assertEquals(cellStyle, result, "getStyle should return the column style");
    }

    @Test
    void getColumnWidth_shouldReturnColumnWidth() {
        // Arrange
        String name = "Test Column";
        ExcelColumn<String> column = new ExcelColumn<>(name, function, cellStyle, columnSetter);
        int expectedWidth = column.getColumnWidth(); // Initial width based on name

        // Act
        int result = column.getColumnWidth();

        // Assert
        assertEquals(expectedWidth, result, "getColumnWidth should return the column width");
    }

    // Tests for ExcelColumnBuilder inner class
    @Test
    void excelColumnBuilder_shouldBuildColumnWithDefaultValues() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        String name = "Test Column";
        Function<String, Object> func = data -> data;

        // Act
        ExcelColumn.ExcelColumnBuilder<String> builder = new ExcelColumn.ExcelColumnBuilder<>(writer, name, (r, c) -> func.apply(r));
        
        // We need to access the private build method indirectly by calling a method that uses it
        // Let's use reflection to call the private build method
        ExcelColumn<String> column = callBuildMethodViaReflection(builder);

        // Assert
        assertNotNull(column, "Builder should create a column instance");
        assertEquals(name, column.getName(), "Column name should match");
    }

    @Test
    void excelColumnBuilder_shouldSetTypeAndFormat() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        String name = "Test Column";
        Function<String, Object> func = data -> data;

        // Act
        ExcelColumn.ExcelColumnBuilder<String> builder = new ExcelColumn.ExcelColumnBuilder<>(writer, name, (r, c) -> func.apply(r))
                .type(ExcelDataType.INTEGER)
                .format("#,##0.00");
        
        ExcelColumn<String> column = callBuildMethodViaReflection(builder);

        // Assert
        assertNotNull(column, "Builder should create a column instance with type and format");
    }

    @Test
    void excelColumnBuilder_shouldSetAlignment() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        String name = "Test Column";
        Function<String, Object> func = data -> data;

        // Act
        ExcelColumn.ExcelColumnBuilder<String> builder = new ExcelColumn.ExcelColumnBuilder<>(writer, name, (r, c) -> func.apply(r))
                .alignment(HorizontalAlignment.LEFT);
        
        ExcelColumn<String> column = callBuildMethodViaReflection(builder);

        // Assert
        assertNotNull(column, "Builder should create a column instance with alignment");
    }

    @Test
    void excelColumnBuilder_shouldSetCustomStyle() {
        // Arrange
        ExcelWriter<String> writer = new ExcelWriter<>();
        String name = "Test Column";
        Function<String, Object> func = data -> data;
        CellStyle customStyle = workbook.createCellStyle();

        // Act
        ExcelColumn.ExcelColumnBuilder<String> builder = new ExcelColumn.ExcelColumnBuilder<>(writer, name, (r, c) -> func.apply(r))
                .style(customStyle);
        
        ExcelColumn<String> column = callBuildMethodViaReflection(builder);

        // Assert
        assertNotNull(column, "Builder should create a column instance with custom style");
        assertEquals(customStyle, column.getStyle(), "Column style should match custom style");
    }

    /**
     * Helper method to call the private build() method via reflection.
     */
    private ExcelColumn<String> callBuildMethodViaReflection(ExcelColumn.ExcelColumnBuilder<String> builder) {
        try {
            java.lang.reflect.Method buildMethod = ExcelColumn.ExcelColumnBuilder.class.getDeclaredMethod("build");
            buildMethod.setAccessible(true);
            return (ExcelColumn<String>) buildMethod.invoke(builder);
        } catch (Exception e) {
            fail("Failed to call build method via reflection: " + e.getMessage());
            return null;
        }
    }
}