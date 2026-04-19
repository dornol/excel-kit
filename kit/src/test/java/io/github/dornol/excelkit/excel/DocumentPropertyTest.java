package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class DocumentPropertyTest {

    @Nested
    class ExcelWriterProperties {

        @Test
        void coreProperties_titleAndAuthor() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .documentProperty("title", "Sales Report")
                    .documentProperty("author", "Finance Team")
                    .documentProperty("subject", "Q4 Sales")
                    .documentProperty("keywords", "sales,revenue")
                    .documentProperty("description", "Quarterly report")
                    .documentProperty("category", "Financial")
                    .column("Name", s -> s)
                    .write(Stream.of("Alice"))
                    .writeTo(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var core = wb.getProperties().getCoreProperties();
                assertEquals("Sales Report", core.getTitle());
                assertEquals("Finance Team", core.getCreator());
                assertEquals("Q4 Sales", core.getSubject());
                assertEquals("sales,revenue", core.getKeywords());
                assertEquals("Quarterly report", core.getDescription());
                assertEquals("Financial", core.getCategory());
            }
        }

        @Test
        void customProperty() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .documentProperty("department", "Engineering")
                    .documentProperty("version", "2.0")
                    .column("Name", s -> s)
                    .write(Stream.of("Alice"))
                    .writeTo(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var custom = wb.getProperties().getCustomProperties();
                assertEquals("Engineering", custom.getProperty("department").getLpwstr());
                assertEquals("2.0", custom.getProperty("version").getLpwstr());
            }
        }

        @Test
        void corePropertyKeys_areCaseInsensitive() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .documentProperty("Title", "Report")
                    .documentProperty("AUTHOR", "Admin")
                    .column("Name", s -> s)
                    .write(Stream.of("Alice"))
                    .writeTo(out);

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var core = wb.getProperties().getCoreProperties();
                assertEquals("Report", core.getTitle());
                assertEquals("Admin", core.getCreator());
            }
        }
    }

    @Nested
    class ExcelWorkbookProperties {

        @Test
        void documentProperty_onWorkbook() throws Exception {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = ExcelWorkbook.create()) {
                wb.documentProperty("title", "Multi-Sheet Report")
                  .documentProperty("author", "System");
                wb.<String>sheet("Data").column("Name", s -> s).write(Stream.of("Bob"));
                wb.finish().writeTo(out);
            }

            try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var core = wb.getProperties().getCoreProperties();
                assertEquals("Multi-Sheet Report", core.getTitle());
                assertEquals("System", core.getCreator());
            }
        }
    }
}
