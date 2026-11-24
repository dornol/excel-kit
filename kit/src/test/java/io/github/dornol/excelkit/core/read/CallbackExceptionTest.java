package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.read.ExcelReadException;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.nio.charset.StandardCharsets;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.*;

/** What the caller's own callback throws reaches the caller as it was thrown, not as a failure of the read. */
class CallbackExceptionTest {

    private static final class Abort extends RuntimeException {
        private static final long serialVersionUID = 1L;
    }

    private static InputStream csv() {
        return new ByteArrayInputStream("A\n1\n2\n".getBytes(StandardCharsets.UTF_8));
    }

    private static InputStream xlsx() throws Exception {
        var out = new ByteArrayOutputStream();
        try (var wb = new XSSFWorkbook()) {
            var sheet = wb.createSheet("S");
            sheet.createRow(0).createCell(0).setCellValue("A");
            sheet.createRow(1).createCell(0).setCellValue("1");
            wb.write(out);
        }
        return new ByteArrayInputStream(out.toByteArray());
    }

    @Nested
    class Csv {
        @Test
        void aResultCallbackExceptionIsThrownAsItWas() {
            assertThrows(Abort.class, () -> CsvReader.forMap().read(csv(), r -> { throw new Abort(); }));
        }

        @Test
        void anOnErrorExceptionIsThrownAsItWas() {
            assertThrows(Abort.class, () -> CsvReader.<String[]>setter(() -> new String[1])
                    .column("A", (p, c) -> { throw new IllegalStateException("bad cell"); })
                    .read(csv(), row -> { }, error -> { throw new Abort(); }));
        }

        @Test
        void anOnSuccessExceptionIsThrownAsItWas() {
            assertThrows(Abort.class, () -> CsvReader.forMap().read(csv(), row -> { throw new Abort(); }, error -> { }));
        }

        @Test
        void aReadStrictConsumerExceptionIsThrownAsItWas() {
            assertThrows(Abort.class, () -> CsvReader.forMap().readStrict(csv(), row -> { throw new Abort(); }));
        }

        @Test
        void aFailureOfTheReadItselfIsStillACsvReadException() {
            assertThrows(CsvReadException.class, () -> CsvReader.forMap("Missing").strictHeaders().read(csv(), r -> { }));
        }
    }

    @Nested
    class Excel {
        @Test
        void aResultCallbackExceptionIsThrownAsItWas() {
            assertThrows(Abort.class, () -> ExcelReader.forMap().read(xlsx(), r -> { throw new Abort(); }));
        }

        @Test
        void anOnSuccessExceptionIsThrownAsItWas() {
            assertThrows(Abort.class, () -> ExcelReader.forMap().read(xlsx(), row -> { throw new Abort(); }, error -> { }));
        }

        @Test
        void aReadStrictConsumerExceptionIsThrownAsItWas() {
            assertThrows(Abort.class, () -> ExcelReader.forMap().readStrict(xlsx(), row -> { throw new Abort(); }));
        }

        @Test
        void aFailureOfTheReadItselfIsStillAnExcelReadException() {
            assertThrows(ExcelReadException.class, () -> ExcelReader.forMap("Missing").strictHeaders().read(xlsx(), r -> { }));
        }
    }

    @Test
    void theErrorLimitIsStillAReadAbort() {
        assertThrows(ReadAbortException.class, () -> CsvReader.<Map<String, String>>mapping(row -> {
            throw new IllegalStateException("never maps");
        }).maxErrors(0).read(csv(), r -> { }));
    }
}
