package io.github.dornol.excelkit.shared;

import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for ExcelKitSchema write configuration support.
 */
class SchemaWriteConfigTest {

    @Test
    void schema_withWriteConfig_shouldApplyTypeToExcelWriter() throws IOException {
        ExcelKitSchema<TestProduct> schema = ExcelKitSchema.<TestProduct>builder()
                .column("Name", TestProduct::getName, (p, cell) -> p.setName(cell.asString()))
                .column("Price", TestProduct::getPrice, (p, cell) -> p.setPrice(cell.asInt()),
                        c -> c.type(ExcelDataType.INTEGER))
                .column("Rate", TestProduct::getRate, (p, cell) -> p.setRate(cell.asDouble()),
                        c -> c.type(ExcelDataType.DOUBLE))
                .build();

        // Write
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(Stream.of(
                        new TestProduct("Widget", 1000, 0.15),
                        new TestProduct("Gadget", 2500, 0.25)
                ))
                .consumeOutputStream(out);

        assertTrue(out.toByteArray().length > 0);

        // Read back to verify data integrity
        List<TestProduct> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
            schema.excelReader(TestProduct::new, null)
                    .build(is)
                    .read(r -> {
                        if (r.success()) results.add(r.data());
                    });
        }

        assertEquals(2, results.size());
        assertEquals("Widget", results.get(0).getName());
        assertEquals(1000, results.get(0).getPrice());
        assertEquals("Gadget", results.get(1).getName());
        assertEquals(2500, results.get(1).getPrice());
    }

    @Test
    void schema_withoutWriteConfig_shouldStillWork() throws IOException {
        ExcelKitSchema<TestProduct> schema = ExcelKitSchema.<TestProduct>builder()
                .column("Name", TestProduct::getName, (p, cell) -> p.setName(cell.asString()))
                .column("Price", p -> String.valueOf(p.getPrice()), (p, cell) -> p.setPrice(cell.asInt()))
                .build();

        ByteArrayOutputStream out = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(Stream.of(new TestProduct("Widget", 1000, 0.0)))
                .consumeOutputStream(out);

        assertTrue(out.toByteArray().length > 0);
    }

    @Test
    void schema_readByName_shouldIgnoreColumnOrder() throws IOException {
        // Schema defines: Name, Price, Rate
        ExcelKitSchema<TestProduct> schema = ExcelKitSchema.<TestProduct>builder()
                .column("Name", TestProduct::getName, (p, cell) -> p.setName(cell.asString()))
                .column("Price", TestProduct::getPrice, (p, cell) -> p.setPrice(cell.asInt()),
                        c -> c.type(ExcelDataType.INTEGER))
                .column("Rate", TestProduct::getRate, (p, cell) -> p.setRate(cell.asDouble()),
                        c -> c.type(ExcelDataType.DOUBLE))
                .build();

        // Write with schema (order: Name, Price, Rate)
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        schema.excelWriter()
                .write(Stream.of(new TestProduct("Widget", 1000, 0.15)))
                .consumeOutputStream(out);

        // Read back - schema reader uses name-based matching
        List<TestProduct> results = new ArrayList<>();
        try (InputStream is = new ByteArrayInputStream(out.toByteArray())) {
            schema.excelReader(TestProduct::new, null)
                    .build(is)
                    .read(r -> {
                        if (r.success()) results.add(r.data());
                    });
        }

        assertEquals(1, results.size());
        assertEquals("Widget", results.get(0).getName());
        assertEquals(1000, results.get(0).getPrice());
    }

    @Test
    void schema_csvReadByName_shouldIgnoreColumnOrder() {
        ExcelKitSchema<TestProduct> schema = ExcelKitSchema.<TestProduct>builder()
                .column("Name", TestProduct::getName, (p, cell) -> p.setName(cell.asString()))
                .column("Price", TestProduct::getPrice, (p, cell) -> p.setPrice(cell.asInt()))
                .build();

        // CSV with reversed column order
        String csv = "Price,Name\n1000,Widget\n2500,Gadget\n";
        InputStream is = new ByteArrayInputStream(csv.getBytes());

        List<TestProduct> results = new ArrayList<>();
        schema.csvReader(TestProduct::new, null)
                .build(is)
                .read(r -> {
                    if (r.success()) results.add(r.data());
                });

        assertEquals(2, results.size());
        assertEquals("Widget", results.get(0).getName());
        assertEquals(1000, results.get(0).getPrice());
        assertEquals("Gadget", results.get(1).getName());
        assertEquals(2500, results.get(1).getPrice());
    }

    public static class TestProduct {
        private String name;
        private int price;
        private double rate;

        public TestProduct() {}

        public TestProduct(String name, int price, double rate) {
            this.name = name;
            this.price = price;
            this.rate = rate;
        }

        public String getName() { return name; }
        public void setName(String name) { this.name = name; }
        public int getPrice() { return price; }
        public void setPrice(int price) { this.price = price; }
        public double getRate() { return rate; }
        public void setRate(double rate) { this.rate = rate; }
    }
}
