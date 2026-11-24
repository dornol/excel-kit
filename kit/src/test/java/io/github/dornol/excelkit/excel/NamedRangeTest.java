package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class NamedRangeTest {

    @Test
    void namedRange_withReferenceString_verifyNameAndFormula() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Category", s -> s)
                .afterData(ctx -> {
                    ctx.namedRange("Categories", "Sheet!$A$2:$A$4");
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("A", "B", "C"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var name = wb.getName("Categories");
            assertNotNull(name, "Named range 'Categories' should exist");
            assertEquals("Sheet!$A$2:$A$4", name.getRefersToFormula());
        }
    }

    @Test
    void namedRange_withColumnIndices_generatesCorrectReference() throws Exception {
        var handler = ExcelWriter.<String>create()
                .sheetName("Data")
                .column("Items", s -> s)
                .afterData(ctx -> {
                    ctx.namedRange("ItemList", 0, 1, 3); // col A, rows 2-4 (0-based: 1-3)
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("X", "Y", "Z"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var name = wb.getName("ItemList");
            assertNotNull(name, "Named range 'ItemList' should exist");
            assertEquals("'Data'!$A$2:$A$4", name.getRefersToFormula());
        }
    }

    @Test
    void namedRange_withApostropheInSheetName_escapesQuote() throws Exception {
        var handler = ExcelWriter.<String>create()
                .sheetName("Bob's data")
                .column("Items", s -> s)
                .namedRange("Fluent", 0)
                .afterData(ctx -> {
                    ctx.namedRange("ViaContext", 0, 1, 3);
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("X", "Y", "Z"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            assertEquals("'Bob''s data'!$A$2:$A$4", wb.getName("Fluent").getRefersToFormula());
            assertEquals("'Bob''s data'!$A$2:$A$4", wb.getName("ViaContext").getRefersToFormula());
        }
    }

    @Test
    void multipleNamedRanges_allCreated() throws Exception {
        var handler = ExcelWriter.<String>create()
                .sheetName("Ref")
                .column("Col", s -> s)
                .afterData(ctx -> {
                    ctx.namedRange("Range1", "'Ref'!$A$2:$A$3");
                    ctx.namedRange("Range2", "'Ref'!$A$2:$A$4");
                    ctx.namedRange("Range3", 0, 1, 3);
                    return ctx.getCurrentRow();
                })
                .write(Stream.of("A", "B", "C"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            assertNotNull(wb.getName("Range1"), "Range1 should exist");
            assertNotNull(wb.getName("Range2"), "Range2 should exist");
            assertNotNull(wb.getName("Range3"), "Range3 should exist");
            assertEquals("'Ref'!$A$2:$A$3", wb.getName("Range1").getRefersToFormula());
            assertEquals("'Ref'!$A$2:$A$4", wb.getName("Range2").getRefersToFormula());
            assertEquals("'Ref'!$A$2:$A$4", wb.getName("Range3").getRefersToFormula());
        }
    }

    @Test
    void namedRange_inExcelWorkbookMultiSheet_verifyAcrossSheets() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Categories")
                    .column("Cat", s -> s)
                    .afterData(ctx -> {
                        ctx.namedRange("CatList", 0, 1, 3);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("Cat1", "Cat2", "Cat3"));

            workbook.<String>sheet("Products")
                    .column("Product", s -> s)
                    .afterData(ctx -> {
                        ctx.namedRange("ProdList", 0, 1, 2);
                        return ctx.getCurrentRow();
                    })
                    .write(Stream.of("Prod1", "Prod2"));

            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var catName = wb.getName("CatList");
            assertNotNull(catName, "CatList named range should exist");
            assertEquals("'Categories'!$A$2:$A$4", catName.getRefersToFormula());

            var prodName = wb.getName("ProdList");
            assertNotNull(prodName, "ProdList named range should exist");
            assertEquals("'Products'!$A$2:$A$3", prodName.getRefersToFormula());
        }
    }
}
