package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.ExcelReadException;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

class ReadHeaderOptionsTest {

    @Test
    void csvReader_columnAliases_resolveFirstMatchingHeader() {
        String csv = "이름,Age\nAlice,30\n";
        List<Person> results = new ArrayList<>();

        CsvReader.setter(Person::new)
                .column(List.of("Name", "이름"), (p, c) -> p.name = c.asString())
                .column("Age", (p, c) -> p.age = c.asInt())
                .read(csv(csv), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("Alice", results.get(0).name);
        assertEquals(30, results.get(0).age);
    }

    @Test
    void csvReader_strictHeaders_failsWhenPositionalColumnHasNoHeader() {
        String csv = "Name\nAlice\n";

        CsvReadException ex = assertThrows(CsvReadException.class, () ->
                CsvReader.setter(Person::new)
                        .strictHeaders()
                        .column((p, c) -> p.name = c.asString())
                        .column((p, c) -> p.age = c.asInt())
                        .read(csv(csv), r -> {}));

        assertTrue(rootMessage(ex).contains("Column index 1 has no header"));
    }

    @Test
    void csvReader_duplicateHeaderPolicy_canUseLastHeader() {
        String csv = "Name,Name\nfirst,last\n";
        List<Person> results = new ArrayList<>();

        CsvReader.setter(Person::new)
                .duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST)
                .column("Name", (p, c) -> p.name = c.asString())
                .read(csv(csv), r -> results.add(r.data()));

        assertEquals("last", results.get(0).name);
    }

    @Test
    void csvReader_duplicateHeaderPolicyFail_failsBeforeDataRows() {
        String csv = "Name,Name\nfirst,last\n";

        CsvReadException ex = assertThrows(CsvReadException.class, () ->
                CsvReader.forMap()
                        .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
                        .read(csv(csv), r -> {}));

        assertTrue(rootMessage(ex).contains("Duplicate header 'Name'"));
    }

    @Test
    void csvReader_forMap_usesConfiguredDuplicateHeaderPolicy() {
        String csv = "Name,Name\nfirst,last\n";
        List<Map<String, String>> results = new ArrayList<>();

        CsvReader.forMap()
                .duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST)
                .read(csv(csv), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals(Map.of("Name", "last"), results.get(0));
    }

    @Test
    void csvReader_rowErrorIncludesPhysicalFileRowNumber() {
        String csv = "Report\nGenerated\nName\n\nAlice\n";
        List<RowError> errors = new ArrayList<>();

        CsvReader.setter(Person::new)
                .headerRowIndex(2)
                .column("Name", (p, c) -> p.name = c.asString())
                .required()
                .read(csv(csv), p -> {}, errors::add);

        assertEquals(1, errors.size());
        assertEquals(1L, errors.get(0).rowNum());
        assertEquals(4L, errors.get(0).fileRowNum());
    }

    @Test
    void csvReader_rowErrorIncludesStructuredCellErrors() {
        String csv = "Name,Age\nAlice,not-a-number\n";
        List<RowError> errors = new ArrayList<>();

        CsvReader.setter(Person::new)
                .column("Name", (p, c) -> p.name = c.asString())
                .column("Age", (p, c) -> p.age = c.asInt())
                .read(csv(csv), p -> {}, errors::add);

        assertEquals(1, errors.size());
        assertEquals(1, errors.get(0).cellErrors().size());
        CellError error = errors.get(0).cellErrors().get(0);
        assertEquals(1, error.columnIndex());
        assertEquals("Age", error.headerName());
        assertEquals("not-a-number", error.cellValue());
        assertTrue(error.message().contains("Failed to set column 'Age'"));
    }

    @Test
    void excelReader_rowErrorIncludesStructuredCellErrors() throws Exception {
        byte[] workbook = workbook(
                List.of("Name", "Age"),
                List.of("Alice", "not-a-number")
        );
        List<RowError> errors = new ArrayList<>();

        ExcelReader.setter(Person::new)
                .column("Name", (p, c) -> p.name = c.asString())
                .column("Age", (p, c) -> p.age = c.asInt())
                .read(new ByteArrayInputStream(workbook), p -> {}, errors::add);

        assertEquals(1, errors.size());
        assertEquals(1, errors.get(0).cellErrors().size());
        CellError error = errors.get(0).cellErrors().get(0);
        assertEquals(1, error.columnIndex());
        assertEquals("Age", error.headerName());
        assertEquals("not-a-number", error.cellValue());
        assertTrue(error.message().contains("Failed to set column 'Age'"));
    }

    @Test
    void csvReader_forMapStrictHeaders_failsWhenSelectedHeaderMissing() {
        String csv = "Name\nAlice\n";

        CsvReadException ex = assertThrows(CsvReadException.class, () ->
                CsvReader.forMap("Name", "Age")
                        .strictHeaders()
                        .read(csv(csv), r -> {}));

        assertTrue(rootMessage(ex).contains("Selected headers [Age] not found"));
    }

    @Test
    void excelReader_columnAliasesAndDuplicatePolicyWorkTogether() throws Exception {
        byte[] workbook = workbook(
                List.of("이름", "Name"),
                List.of("first", "last")
        );
        List<Person> results = new ArrayList<>();

        ExcelReader.setter(Person::new)
                .duplicateHeaderPolicy(DuplicateHeaderPolicy.LAST)
                .column(List.of("Name", "이름"), (p, c) -> p.name = c.asString())
                .read(new ByteArrayInputStream(workbook), r -> results.add(r.data()));

        assertEquals(1, results.size());
        assertEquals("last", results.get(0).name);
    }

    @Test
    void excelReader_duplicateHeaderPolicyFail_failsBeforeDataRows() throws Exception {
        byte[] workbook = workbook(
                List.of("Name", "Name"),
                List.of("first", "last")
        );

        ExcelReadException ex = assertThrows(ExcelReadException.class, () ->
                ExcelReader.forMap()
                        .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
                        .read(new ByteArrayInputStream(workbook), r -> {}));

        assertTrue(rootMessage(ex).contains("Duplicate header 'Name'"));
    }

    @Test
    void excelReader_rowErrorIncludesPhysicalFileRowNumber() throws Exception {
        byte[] workbook = workbookWithHeaderOffset();
        List<RowError> errors = new ArrayList<>();

        ExcelReader.setter(Person::new)
                .headerRowIndex(2)
                .column("Name", (p, c) -> p.name = c.asString())
                .required()
                .read(new ByteArrayInputStream(workbook), p -> {}, errors::add);

        assertEquals(1, errors.size());
        assertEquals(1L, errors.get(0).rowNum());
        assertEquals(4L, errors.get(0).fileRowNum());
    }

    @Test
    void excelReader_forMapStrictHeaders_failsWhenSelectedHeaderMissing() throws Exception {
        byte[] workbook = workbook(
                List.of("Name"),
                List.of("Alice")
        );

        ExcelReadException ex = assertThrows(ExcelReadException.class, () ->
                ExcelReader.forMap("Name", "Age")
                        .strictHeaders()
                        .read(new ByteArrayInputStream(workbook), r -> {}));

        assertTrue(rootMessage(ex).contains("Selected headers [Age] not found"));
    }

    private static ByteArrayInputStream csv(String content) {
        return new ByteArrayInputStream(content.getBytes(StandardCharsets.UTF_8));
    }

    private static byte[] workbook(List<String> headers, List<String> values) throws Exception {
        try (Workbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("Data");
            Row header = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                header.createCell(i).setCellValue(headers.get(i));
            }
            Row data = sheet.createRow(1);
            for (int i = 0; i < values.size(); i++) {
                data.createCell(i).setCellValue(values.get(i));
            }
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] workbookWithHeaderOffset() throws Exception {
        try (Workbook wb = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = wb.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("Report");
            sheet.createRow(1).createCell(0).setCellValue("Generated");
            sheet.createRow(2).createCell(0).setCellValue("Name");
            sheet.createRow(3).createCell(0).setCellValue("");
            sheet.createRow(4).createCell(0).setCellValue("Alice");
            wb.write(out);
            return out.toByteArray();
        }
    }

    private static String rootMessage(Throwable throwable) {
        Throwable current = throwable;
        while (current.getCause() != null) {
            current = current.getCause();
        }
        return current.getMessage();
    }

    static class Person {
        String name;
        int age;
    }
}
