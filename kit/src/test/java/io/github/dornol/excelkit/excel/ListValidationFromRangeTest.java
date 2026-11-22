package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ListValidationFromRangeTest {

    @Test
    void listFromRange_createsValidationOnSheet() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Category", s -> s, c -> c
                        .validation(ExcelValidation.listFromRange("Sheet2!$A$1:$A$5")))
                .write(Stream.of("Test"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertFalse(validations.isEmpty(),
                    "Data validations should exist on the sheet");
        }
    }

    @Test
    void listFromRange_withErrorConfiguration_verifyErrorBox() throws Exception {
        var handler = ExcelWriter.<String>create()
                .column("Status", s -> s, c -> c
                        .validation(ExcelValidation.listFromRange("Statuses!$A$1:$A$3")
                                .errorTitle("Invalid Status")
                                .errorMessage("Please select a valid status.")
                                .showError(true)))
                .write(Stream.of("Active"));

        var baos = new ByteArrayOutputStream();
        handler.writeTo(baos);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertFalse(validations.isEmpty());
            var dv = validations.get(0);
            assertTrue(dv.getShowErrorBox(),
                    "Error box should be shown");
        }
    }

    @Test
    void listFromRange_factoryMethodReturnsNonNull() {
        var v = ExcelValidation.listFromRange("Options!$B$1:$B$10");
        assertNotNull(v, "listFromRange should return a non-null ExcelValidation");
    }

    @Test
    void listFromRange_combinedWithExcelSheetWriter_validationApplied() throws Exception {
        var baos = new ByteArrayOutputStream();
        try (var workbook = ExcelWorkbook.create()) {
            workbook.<String>sheet("Data")
                    .column("Type", s -> s, c -> c
                            .validation(ExcelValidation.listFromRange("Lookup!$A$1:$A$5")))
                    .write(Stream.of("Item1"));
            workbook.finish().writeTo(baos);
        }

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(baos.toByteArray()))) {
            var validations = wb.getSheetAt(0).getDataValidations();
            assertFalse(validations.isEmpty(),
                    "Data validations should be applied via ExcelSheetWriter");
        }
    }
}
