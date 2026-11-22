package io.github.dornol.excelkit.spring;

import io.github.dornol.excelkit.core.read.CellError;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.read.RowError;
import org.junit.jupiter.api.Test;

import java.util.Arrays;
import java.util.List;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

/** The JSON-friendly form of a row error: defaults for what is missing, and the type taken from the result. */
class UploadErrorTest {

    @Test
    void whatIsMissingBecomesEmptyAndWhatIsGivenIsCopied() {
        UploadError error = new UploadError(3, 5, RowError.Type.VALIDATION, null, null, null);

        assertEquals(List.of(), error.messages());
        assertEquals(List.of(), error.cellErrors());
        assertEquals(List.of(), error.rawValues());

        List<String> messages = new java.util.ArrayList<>(List.of("bad"));
        UploadError copy = new UploadError(3, 5, RowError.Type.VALIDATION, messages, List.of(), List.of("x"));
        messages.add("changed later");

        assertEquals(List.of("bad"), copy.messages(), "the error does not follow a list the caller keeps changing");
        assertThrows(UnsupportedOperationException.class, () -> copy.messages().add("more"));
    }

    @Test
    void anErrorWithoutRawValuesHasNone() {
        UploadError error = new UploadError(1, 2, RowError.Type.MAPPING, List.of("m"),
                List.of(new CellErrorResponse(0, "Name", "v", "bad")));

        assertEquals(List.of(), error.rawValues());
        assertEquals(1, error.cellErrors().size());
        assertEquals("Name", error.cellErrors().get(0).headerName());
    }

    @Test
    void aRowErrorIsConvertedWithItsPositionsCellsAndRawValues() {
        RowError row = new RowError(4, 9, RowError.Type.MAPPING, List.of("could not convert"), new IllegalStateException("x"),
                List.of(new CellError(1, "Age", "abc", "not a number")), List.of("Ann", "abc"));

        UploadError error = UploadError.from(row);

        assertEquals(4, error.rowNum());
        assertEquals(9, error.fileRowNum());
        assertEquals(RowError.Type.MAPPING, error.type());
        assertEquals(List.of("could not convert"), error.messages());
        assertEquals(List.of("Ann", "abc"), error.rawValues());
        assertEquals(1, error.cellErrors().size());
        assertEquals("Age", error.cellErrors().get(0).headerName());
        assertEquals("abc", error.cellErrors().get(0).cellValue());
    }

    @Test
    void aFailedResultKeepsItsOwnErrorType() {
        ReadResult<String> validation = ReadResult.validationFailure("d", List.of("blank"), 7, List.of(), List.of("r"));
        ReadResult<String> mapping = ReadResult.mappingFailure("d", List.of("convert"), new RuntimeException("c"), 8,
                List.of(new CellError(0, "H", "v", "m")), List.of());

        UploadError fromValidation = UploadError.from(2, validation);
        UploadError fromMapping = UploadError.from(3, mapping);

        assertEquals(RowError.Type.VALIDATION, fromValidation.type());
        assertEquals(2, fromValidation.rowNum());
        assertEquals(7, fromValidation.fileRowNum());
        assertEquals(List.of("r"), fromValidation.rawValues());
        assertEquals(RowError.Type.MAPPING, fromMapping.type());
        assertEquals(1, fromMapping.cellErrors().size());
    }

    @Test
    void aResultWithoutAnErrorTypeIsTakenForAValidationErrorUnlessItHasACause() {
        // a successful result has no error type; it is what a caller converting every result would meet
        ReadResult<String> success = ReadResult.success("ok");
        ReadResult<String> withCause = new ReadResult<>("d", false, null, List.of("x"), new RuntimeException("boom"), 1,
                List.of(), List.of());

        assertEquals(RowError.Type.VALIDATION, UploadError.from(1, success).type());
        assertEquals(RowError.Type.MAPPING, UploadError.from(1, withCause).type(),
                "a failed result with a cause and no type is a mapping failure");
        assertTrue(UploadError.from(1, success).messages().isEmpty());
        assertEquals(Arrays.asList(), UploadError.from(1, success).rawValues());
    }
}
