package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.*;

import java.io.*;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("Bugfix: required column in sparse Excel row")
class RequiredSparseRowTest {

    @Test
    void required_sparseRow_missingColumn_shouldFail() throws Exception {
        // Write Excel with 3 columns, but third column is null (SAX won't emit it)
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter<String[]> writer = ExcelWriter.<String[]>create();
        writer.column("Name", arr -> arr[0])
                .column("Age", arr -> arr[1])
                .column("City", arr -> arr[2]);
        writer.write(Stream.<String[]>of(new String[]{"Alice", "30", null})).writeTo(bos);

        // Read with required() on the third column (which is empty/missing)
        List<ReadResult<String[]>> results = new ArrayList<>();
        ExcelReader.setter(() -> new String[3])
                .column("Name", (arr, c) -> arr[0] = c.asString())
                .column("Age", (arr, c) -> arr[1] = c.asString())
                .column("City", (arr, c) -> arr[2] = c.asString()).required()
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertFalse(results.get(0).success(), "Required column with missing cell in sparse row should fail");
        assertTrue(results.get(0).messages().stream().anyMatch(m -> m.contains("Required")));
    }

    @Test
    void nonRequired_sparseRow_missingColumn_shouldSucceed() throws Exception {
        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        ExcelWriter<String[]> writer = ExcelWriter.<String[]>create();
        writer.column("Name", arr -> arr[0])
                .column("Age", arr -> arr[1])
                .column("City", arr -> arr[2]);
        writer.write(Stream.<String[]>of(new String[]{"Alice", "30", null})).writeTo(bos);

        List<ReadResult<String[]>> results = new ArrayList<>();
        ExcelReader.setter(() -> new String[3])
                .column("Name", (arr, c) -> arr[0] = c.asString())
                .column("Age", (arr, c) -> arr[1] = c.asString())
                .column("City", (arr, c) -> arr[2] = c.asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success(), "Non-required missing column should succeed");
    }
}
