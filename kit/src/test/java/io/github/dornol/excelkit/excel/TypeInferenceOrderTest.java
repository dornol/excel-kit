package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;

/**
 * Pins how an untyped column infers its type and number format: the first non-null value picks the column's
 * type and default format; values of another type keep their own type, and numbers of another kind stay
 * numeric. The stored value is always right, only the display format depends on the order of the data.
 */
class TypeInferenceOrderTest {

    private record Shown(CellType type, String format) {
        @Override
        public String toString() {
            return type + "[" + format + "]";
        }
    }

    /** Type and format of the data cells of an untyped column; a null value is left out. */
    private static List<Shown> write(Object... values) throws Exception {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Object>create().column("V", v -> v).write(Arrays.asList(values)).writeTo(out);
        List<Shown> shown = new ArrayList<>();
        try (XSSFWorkbook wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            Sheet sheet = wb.getSheetAt(0);
            for (int r = 1; r <= sheet.getLastRowNum(); r++) {
                if (values[r - 1] == null) continue;
                Cell cell = sheet.getRow(r).getCell(0);
                shown.add(new Shown(cell.getCellType(), cell.getCellStyle().getDataFormatString()));
            }
        }
        return shown;
    }

    private static Shown num(String format) {
        return new Shown(CellType.NUMERIC, format);
    }

    @Nested
    class LeadingNull {
        @Test
        void nullFirst_typeComesFromTheFirstNonNullValue() throws Exception {
            assertEquals(List.of(num("#,##0"), num("#,##0")), write(null, 1, 2));
        }

        @Test
        void nullBetweenValues_doesNotChangeTheFormat() throws Exception {
            assertEquals(List.of(num("#,##0"), num("#,##0")), write(1, null, 3));
        }
    }

    @Nested
    class MixedNumbers {
        @Test
        void integerThenDecimal_decimalKeepsGeneralFormat() throws Exception {
            assertEquals(List.of(num("#,##0"), num("General")), write(1, 2.5));
        }

        @Test
        void decimalThenInteger_bothUseTheDecimalFormat() throws Exception {
            assertEquals(List.of(num("#,##0.00"), num("#,##0.00")), write(2.5, 1));
        }

        @Test
        void integerThenLong_longKeepsGeneralFormat() throws Exception {
            assertEquals(List.of(num("#,##0"), num("General")), write(1, 5_000_000_000L));
        }

        @Test
        void bigDecimalThenInteger_bothUseTheDecimalFormat() throws Exception {
            assertEquals(List.of(num("#,##0.00"), num("#,##0.00")), write(new BigDecimal("1.5"), 1));
        }
    }

    @Nested
    class MixedKinds {
        @Test
        void numberThenText_eachKeepsItsOwnType() throws Exception {
            assertEquals(List.of(num("#,##0"), new Shown(CellType.STRING, "General")), write(1, "a"));
        }

        @Test
        void textThenNumber_numberStillGetsItsNumberFormat() throws Exception {
            assertEquals(List.of(new Shown(CellType.STRING, "General"), num("#,##0")), write("a", 1));
        }

        @Test
        void dateAndDateTime_eachKeepsItsOwnFormatWhicheverComesFirst() throws Exception {
            LocalDate date = LocalDate.of(2024, 1, 1);
            LocalDateTime dateTime = LocalDateTime.of(2024, 1, 1, 1, 1);
            assertEquals(List.of(num("yyyy-mm-dd"), num("yyyy-mm-dd hh:mm:ss")), write(date, dateTime));
            assertEquals(List.of(num("yyyy-mm-dd hh:mm:ss"), num("yyyy-mm-dd")), write(dateTime, date));
        }
    }
}
