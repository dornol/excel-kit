package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.csv.CsvReadException;
import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.shared.ExcelKitException;
import io.github.dornol.excelkit.shared.TempResourceCreateException;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for exception classes — ensures all constructors work correctly.
 */
class ExceptionTest {

    @Nested
    class ExcelWriteExceptionTests {

        @Test
        void messageOnly_constructor() {
            ExcelWriteException e = new ExcelWriteException("test error");
            assertEquals("test error", e.getMessage());
            assertNull(e.getCause());
        }

        @Test
        void messageAndCause_constructor() {
            RuntimeException cause = new RuntimeException("root cause");
            ExcelWriteException e = new ExcelWriteException("test error", cause);
            assertEquals("test error", e.getMessage());
            assertSame(cause, e.getCause());
        }

        @Test
        void extendsExcelKitException() {
            assertInstanceOf(ExcelKitException.class, new ExcelWriteException("x"));
        }
    }

    @Nested
    class ExcelReadExceptionTests {

        @Test
        void messageOnly_constructor() {
            ExcelReadException e = new ExcelReadException("read error");
            assertEquals("read error", e.getMessage());
            assertNull(e.getCause());
        }

        @Test
        void messageAndCause_constructor() {
            RuntimeException cause = new RuntimeException("io error");
            ExcelReadException e = new ExcelReadException("read error", cause);
            assertEquals("read error", e.getMessage());
            assertSame(cause, e.getCause());
        }

        @Test
        void extendsExcelKitException() {
            assertInstanceOf(ExcelKitException.class, new ExcelReadException("x"));
        }
    }

    @Nested
    class CsvWriteExceptionTests {

        @Test
        void messageOnly_constructor() {
            CsvWriteException e = new CsvWriteException("csv error");
            assertEquals("csv error", e.getMessage());
            assertNull(e.getCause());
        }

        @Test
        void messageAndCause_constructor() {
            RuntimeException cause = new RuntimeException("io");
            CsvWriteException e = new CsvWriteException("csv error", cause);
            assertEquals("csv error", e.getMessage());
            assertSame(cause, e.getCause());
        }

        @Test
        void extendsExcelKitException() {
            assertInstanceOf(ExcelKitException.class, new CsvWriteException("x"));
        }
    }

    @Nested
    class CsvReadExceptionTests {

        @Test
        void messageOnly_constructor() {
            CsvReadException e = new CsvReadException("csv read");
            assertEquals("csv read", e.getMessage());
            assertNull(e.getCause());
        }

        @Test
        void messageAndCause_constructor() {
            RuntimeException cause = new RuntimeException("parse");
            CsvReadException e = new CsvReadException("csv read", cause);
            assertEquals("csv read", e.getMessage());
            assertSame(cause, e.getCause());
        }
    }

    @Nested
    class TempResourceCreateExceptionTests {

        @Test
        void causeOnly_constructor() {
            RuntimeException cause = new RuntimeException("io");
            TempResourceCreateException e = new TempResourceCreateException(cause);
            assertSame(cause, e.getCause());
        }

        @Test
        void extendsExcelKitException() {
            assertInstanceOf(ExcelKitException.class,
                    new TempResourceCreateException(new RuntimeException()));
        }
    }

    @Nested
    class ExcelKitExceptionTests {

        @Test
        void messageOnly() {
            ExcelKitException e = new ExcelKitException("msg");
            assertEquals("msg", e.getMessage());
        }

        @Test
        void messageAndCause() {
            RuntimeException cause = new RuntimeException();
            ExcelKitException e = new ExcelKitException("msg", cause);
            assertEquals("msg", e.getMessage());
            assertSame(cause, e.getCause());
        }

        @Test
        void causeOnly() {
            RuntimeException cause = new RuntimeException("root");
            ExcelKitException e = new ExcelKitException(cause);
            assertSame(cause, e.getCause());
        }

        @Test
        void extendsRuntimeException() {
            assertInstanceOf(RuntimeException.class, new ExcelKitException("x"));
        }
    }
}
