package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.excel.ExcelReadException;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class ExcelKitSchemaReadOptionsTest {

    private ExcelKitSchema<TestPerson> schema;

    @BeforeEach
    void setUp() {
        schema = ExcelKitSchema.<TestPerson>builder()
                .column("Name", TestPerson::getName, (p, cell) -> p.setName(cell.asString()))
                .column("Age", TestPerson::getAge, (p, cell) -> p.setAge(cell.asInt()))
                .build();
    }

    @Test
    void csvReader_shouldUseHeaderAliasesAndRequiredColumns() {
        ExcelKitSchema<TestPerson> aliasSchema = aliasSchema();
        String csv = "Full Name,Age\nAlice,30\n,25\n";
        List<TestPerson> valid = new ArrayList<>();
        List<RowError> errors = new ArrayList<>();

        aliasSchema.csvReader(TestPerson::new, null)
                .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                .read(valid::add, errors::add);

        assertEquals(1, valid.size());
        assertEquals("Alice", valid.get(0).getName());
        assertEquals(1, errors.size());
        assertEquals("Full Name", errors.get(0).cellErrors().get(0).headerName());
    }

    @Test
    void excelReader_shouldUseHeaderAliasesAndRequiredColumns() throws IOException {
        ExcelKitSchema<TestPerson> aliasSchema = aliasSchema();
        byte[] workbook = workbook(
                List.of("Full Name", "Age"),
                List.of(List.of("Alice", "30"), List.of("", "25"))
        );
        List<TestPerson> valid = new ArrayList<>();
        List<RowError> errors = new ArrayList<>();

        aliasSchema.excelReader(TestPerson::new, null)
                .build(new ByteArrayInputStream(workbook))
                .read(valid::add, errors::add);

        assertEquals(1, valid.size());
        assertEquals("Alice", valid.get(0).getName());
        assertEquals(1, errors.size());
        assertEquals("Full Name", errors.get(0).cellErrors().get(0).headerName());
    }

    @Test
    void csvReader_duplicateHeaderPolicyFail_failsBeforeRows() {
        String csv = "Name,Name,Age\nfirst,last,30\n";

        CsvReadException ex = assertThrows(CsvReadException.class, () ->
                schema.csvReader(TestPerson::new, null)
                        .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
                        .build(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)))
                        .read(row -> {}));

        assertTrue(rootMessage(ex).contains("Duplicate header 'Name'"));
    }

    @Test
    void excelReader_duplicateHeaderPolicyFail_failsBeforeRows() throws IOException {
        byte[] workbook = workbook(
                List.of("Name", "Name", "Age"),
                List.of(List.of("first", "last", "30"))
        );

        ExcelReadException ex = assertThrows(ExcelReadException.class, () ->
                schema.excelReader(TestPerson::new, null)
                        .duplicateHeaderPolicy(DuplicateHeaderPolicy.FAIL)
                        .build(new ByteArrayInputStream(workbook))
                        .read(row -> {}));

        assertTrue(rootMessage(ex).contains("Duplicate header 'Name'"));
    }

    private static ExcelKitSchema<TestPerson> aliasSchema() {
        return ExcelKitSchema.<TestPerson>builder()
                .requiredColumn("Name", List.of("Full Name", "이름"),
                        TestPerson::getName, (p, cell) -> p.setName(cell.asString()))
                .column("Age", TestPerson::getAge, (p, cell) -> p.setAge(cell.asInt()))
                .build();
    }

    private static byte[] workbook(List<String> headers, List<List<String>> rows) throws IOException {
        try (Workbook workbook = new XSSFWorkbook(); ByteArrayOutputStream out = new ByteArrayOutputStream()) {
            Sheet sheet = workbook.createSheet("Data");
            Row header = sheet.createRow(0);
            for (int i = 0; i < headers.size(); i++) {
                header.createCell(i).setCellValue(headers.get(i));
            }
            for (int rowIndex = 0; rowIndex < rows.size(); rowIndex++) {
                Row row = sheet.createRow(rowIndex + 1);
                List<String> values = rows.get(rowIndex);
                for (int col = 0; col < values.size(); col++) {
                    row.createCell(col).setCellValue(values.get(col));
                }
            }
            workbook.write(out);
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

    public static class TestPerson {
        private String name;
        private int age;

        public TestPerson() {
        }

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
