package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ConditionalFormattingRulesTest {
    @Test
    void conditionalFormatting_noRules_shouldNotFail() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s)
                .conditionalFormatting(cf -> {})  // no rules added
                .write(Stream.of("data"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(0, wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings());
        }
    }

    @Test
    void conditionalFormatting_allColumnsByDefault() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("A", s -> s)
                .column("B", s -> s)
                .conditionalFormatting(cf -> cf
                        .equalTo("\"test\"", ExcelColor.LIGHT_RED))
                .write(Stream.of("test"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            // Rules applied to all columns
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() > 0);
        }
    }

    @Test
    void conditionalFormatting_multipleRulesSets() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Val", s -> s)
                .conditionalFormatting(cf -> cf
                        .columns(0)
                        .greaterThan("100", ExcelColor.RED))
                .conditionalFormatting(cf -> cf
                        .columns(0)
                        .lessThan("0", ExcelColor.BLUE))
                .write(Stream.of("50"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertTrue(wb.getSheetAt(0).getSheetConditionalFormatting()
                    .getNumConditionalFormattings() >= 2);
        }
    }
}
