package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.ConditionalFormatting;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ConditionalFormattingTest {

    record Product(String name, int price) {}

    @Test
    void conditionalFormatting_greaterThan() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("100", ExcelColor.LIGHT_RED))
                .write(Stream.of(new Product("A", 50), new Product("B", 200)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertTrue(scf.getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_multipleRules() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("200", ExcelColor.LIGHT_RED)
                        .lessThan("50", ExcelColor.LIGHT_GREEN)
                        .between("50", "200", ExcelColor.LIGHT_YELLOW))
                .write(Stream.of(new Product("A", 100)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            SheetConditionalFormatting scf = wb.getSheetAt(0).getSheetConditionalFormatting();
            assertTrue(scf.getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_equalTo() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Name", Product::name)
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .equalTo("100", ExcelColor.LIGHT_BLUE))
                .write(Stream.of(new Product("A", 100)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_notEqualTo() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .notEqualTo("0", ExcelColor.LIGHT_ORANGE))
                .write(Stream.of(new Product("A", 100)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_greaterThanOrEqual() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .greaterThanOrEqual("100", ExcelColor.LIGHT_RED))
                .write(Stream.of(new Product("A", 100)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_lessThanOrEqual() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .lessThanOrEqual("50", ExcelColor.LIGHT_GREEN))
                .write(Stream.of(new Product("A", 30)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_notBetween() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .notBetween("10", "90", ExcelColor.CORAL))
                .write(Stream.of(new Product("A", 100)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_inExcelSheetWriter() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        try (ExcelWorkbook wb = new ExcelWorkbook()) {
            wb.<Product>sheet("Products")
                    .column("Name", Product::name)
                    .column("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                    .conditionalFormatting(cf -> cf
                            .columns(1)
                            .greaterThan("100", ExcelColor.LIGHT_RED))
                    .write(Stream.of(new Product("A", 200)));
            wb.finish().consumeOutputStream(out);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_withStartRow() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<Product>()
                .addColumn("Price", p -> p.price, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf
                        .startRow(2)
                        .greaterThan("100", ExcelColor.LIGHT_RED))
                .write(Stream.of(new Product("A", 200)))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }
}
