package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadColumn;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.core.read.RowError;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("B6: Reader required()")
class RequiredColumnTest {

    @Test
    void required_excelReader_emptyCell_shouldFail() throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter<String[]> writer = ExcelWriter.<String[]>create();
        writer.column("Name", arr -> arr[0])
                .column("Age", arr -> arr[1]);
        writer.write(Stream.<String[]>of(new String[]{"", "30"})).writeTo(bos);

        List<ReadResult<String[]>> results = new ArrayList<>();
        ExcelReader.setter(() -> new String[2])
                .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                .column("Age", (arr, c) -> arr[1] = c.asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertFalse(results.get(0).success(), "Required column with blank value should fail");
        assertEquals(RowError.Type.VALIDATION, results.get(0).errorType());
        assertNotNull(results.get(0).messages());
        assertTrue(results.get(0).messages().stream().anyMatch(m -> m.contains("Required")));
    }

    @Test
    void required_excelReader_nonEmptyCell_shouldSucceed() throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter<String[]> writer = ExcelWriter.<String[]>create();
        writer.column("Name", arr -> arr[0])
                .column("Age", arr -> arr[1]);
        writer.write(Stream.<String[]>of(new String[]{"Alice", "30"})).writeTo(bos);

        List<ReadResult<String[]>> results = new ArrayList<>();
        ExcelReader.setter(() -> new String[2])
                .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                .column("Age", (arr, c) -> arr[1] = c.asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data()[0]);
    }

    @Test
    void required_csvReader_emptyCell_shouldFail() {
        String csv = "Name,Age\n,30";

        List<ReadResult<String[]>> results = new ArrayList<>();
        CsvReader.setter(() -> new String[2])
                .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                .column("Age", (arr, c) -> arr[1] = c.asString())
                .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);

        assertEquals(1, results.size());
        assertFalse(results.get(0).success());
        assertEquals(RowError.Type.VALIDATION, results.get(0).errorType());
        assertTrue(results.get(0).messages().stream().anyMatch(m -> m.contains("Required")));
    }

    @Test
    void required_csvReader_nonEmptyCell_shouldSucceed() {
        String csv = "Name,Age\nBob,25";

        List<ReadResult<String[]>> results = new ArrayList<>();
        CsvReader.setter(() -> new String[2])
                .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                .column("Age", (arr, c) -> arr[1] = c.asString())
                .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Bob", results.get(0).data()[0]);
    }

    @Test
    void required_multipleRequiredColumns_bothEmpty_shouldFail() {
        String csv = "Name,Age\n,";

        List<ReadResult<String[]>> results = new ArrayList<>();
        CsvReader.setter(() -> new String[2])
                .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                .column("Age", (arr, c) -> arr[1] = c.asString()).required()
                .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);

        assertEquals(1, results.size());
        assertFalse(results.get(0).success());
        assertEquals(2, results.get(0).messages().size(), "Both required columns should report errors");
    }

    @Test
    void required_onlySecondEmpty_shouldFailWithColumnName() {
        String csv = "Name,Age\nAlice,";

        List<ReadResult<String[]>> results = new ArrayList<>();
        CsvReader.setter(() -> new String[2])
                .column("Name", (arr, c) -> arr[0] = c.asString()).required()
                .column("Age", (arr, c) -> arr[1] = c.asString()).required()
                .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);

        assertEquals(1, results.size());
        assertFalse(results.get(0).success());
        assertEquals(1, results.get(0).messages().size());
        assertTrue(results.get(0).messages().get(0).contains("Age"));
    }

    @Test
    void required_withoutCalling_shouldNotValidateEmpty() {
        String csv = "Name,Age\n,30";

        List<ReadResult<String[]>> results = new ArrayList<>();
        CsvReader.setter(() -> new String[2])
                .column("Name", (arr, c) -> arr[0] = c.asString())
                .column("Age", (arr, c) -> arr[1] = c.asString())
                .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success(), "Non-required column should pass with blank value");
    }

    @Test
    void required_noColumnRegistered_shouldThrow() {
        assertThrows(IllegalStateException.class, () ->
                CsvReader.setter(() -> new String[1]).required());
    }

    @Test
    void required_noColumnRegistered_excelReader_shouldThrow() {
        assertThrows(IllegalStateException.class, () ->
                ExcelReader.setter(() -> new String[1]).required());
    }

    @Test
    void readColumn_required_shouldReturnNewInstance() {
        ReadColumn<String> col = new ReadColumn<>("Name", (s, c) -> {});
        assertFalse(col.isRequired());

        ReadColumn<String> req = col.required();
        assertTrue(req.isRequired());
        assertFalse(col.isRequired(), "Original should be unchanged");
        assertEquals("Name", req.headerName());
    }
}
