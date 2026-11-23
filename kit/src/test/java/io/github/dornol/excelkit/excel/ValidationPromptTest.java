package io.github.dornol.excelkit.excel;

import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ValidationPromptTest {

    private static org.apache.poi.ss.usermodel.DataValidation write(ExcelValidation validation) throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Age", s -> s, c -> c.validation(validation))
                .write(Stream.of("1"))
                .writeTo(out);
        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            return wb.getSheetAt(0).getDataValidations().get(0);
        }
    }

    @Nested
    class InputMessage {
        @Test
        void showsATooltipWithTitleAndText() throws IOException {
            var v = write(ExcelValidation.integerBetween(1, 150).inputTitle("Age").inputMessage("Enter 1-150"));
            assertTrue(v.getShowPromptBox());
            assertEquals("Age", v.getPromptBoxTitle());
            assertEquals("Enter 1-150", v.getPromptBoxText());
        }

        @Test
        void aMessageAloneIsEnough() throws IOException {
            var v = write(ExcelValidation.integerBetween(1, 150).inputMessage("Enter 1-150"));
            assertTrue(v.getShowPromptBox());
            assertEquals("Enter 1-150", v.getPromptBoxText());
        }

        @Test
        void noTooltipByDefault() throws IOException {
            assertFalse(write(ExcelValidation.integerBetween(1, 150)).getShowPromptBox());
        }

        @Test
        void limitsAreEnforced() {
            var v = ExcelValidation.integerBetween(1, 2);
            assertThrows(IllegalArgumentException.class, () -> v.inputTitle("x".repeat(33)));
            assertThrows(IllegalArgumentException.class, () -> v.inputMessage("x".repeat(256)));
            assertDoesNotThrow(() -> v.inputTitle("x".repeat(32)).inputMessage("x".repeat(255)));
        }
    }

    @Nested
    class AllowBlank {
        @Test
        void blankIsAllowedByDefault() throws IOException {
            assertTrue(write(ExcelValidation.integerBetween(1, 150)).getEmptyCellAllowed());
        }

        @Test
        void canBeTurnedOff() throws IOException {
            assertFalse(write(ExcelValidation.integerBetween(1, 150).allowBlank(false)).getEmptyCellAllowed());
        }
    }
}
