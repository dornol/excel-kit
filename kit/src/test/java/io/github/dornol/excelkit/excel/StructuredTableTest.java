package io.github.dornol.excelkit.excel;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

class StructuredTableTest {

    @Test void writerSupportsStreamingOptionsAndStructuredTable() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create(options -> options.rowAccessWindowSize(1)
                        .compressTempFiles(true).useSharedStrings(true))
                .column("Name", Person::name)
                .table("PeopleTable")
                .write(List.of(new Person("A"), new Person("B"))).writeTo(output);
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("PeopleTable", workbook.getSheetAt(0).getTables().get(0).getName());
        }
    }

    @Test void structuredTableRejectsInvalidAndCellReferenceNames() {
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.create().table("bad name"));
        assertThrows(IllegalArgumentException.class, () -> ExcelWriter.create().table("A1"));
    }

    @Test void rolloverCreatesUniqueStructuredTables() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        ExcelWriter.<Person>create().column("Name", Person::name).maxRows(2)
                .table(new io.github.dornol.excelkit.excel.TableOptions(
                        "People", "TableStyleMedium3", true, true))
                .write(List.of(new Person("A"), new Person("B"), new Person("C"))).writeTo(output);
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals(2, workbook.getNumberOfSheets());
            assertEquals("People_1", workbook.getSheetAt(0).getTables().get(0).getName());
            assertEquals("People_2", workbook.getSheetAt(1).getTables().get(0).getName());
        }
    }

    @Test void workbookSheetSupportsSharedStreamingOptionsAndTables() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var workbook = io.github.dornol.excelkit.excel.ExcelWorkbook.create(options ->
                options.streaming(new io.github.dornol.excelkit.excel.StreamingOptions(1, true, true)))) {
            workbook.<Person>sheet("People").column("Name", Person::name)
                    .table("WorkbookPeople").write(List.of(new Person("A"), new Person("B")));
            workbook.finish().writeTo(output);
        }
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("WorkbookPeople", workbook.getSheet("People").getTables().get(0).getName());
        }
    }

    @Test void templateListSupportsStructuredTable() throws Exception {
        ByteArrayOutputStream template = new ByteArrayOutputStream();
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            workbook.createSheet("Data").createRow(0).createCell(0).setCellValue("Name");
            workbook.write(template);
        }
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var writer = new io.github.dornol.excelkit.excel.ExcelTemplateWriter(
                new ByteArrayInputStream(template.toByteArray()))) {
            writer.<Person>list(1).column("Name", Person::name).table("TemplatePeople")
                    .write(List.of(new Person("A")));
            writer.finish().writeTo(output);
        }
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook(
                new ByteArrayInputStream(output.toByteArray()))) {
            assertEquals("TemplatePeople", workbook.getSheet("Data").getTables().get(0).getName());
        }
    }

    record Person(String name) {}
}
