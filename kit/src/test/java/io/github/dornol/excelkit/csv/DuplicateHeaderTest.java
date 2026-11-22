package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.read.DuplicateHeaderPolicy;
import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.excel.*;
import org.junit.jupiter.api.*;

import java.io.*;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;

import static org.junit.jupiter.api.Assertions.*;

@DisplayName("A5: Duplicate header — mapping mode uses first occurrence")
class DuplicateHeaderTest {

    @Test
    void csv_duplicateHeaders_mappingMode_shouldUseFirstOccurrence() {
        // CSV allows duplicate headers in raw data
        String csv = "Name,Name,Age\nAlice,Bob,30";

        List<ReadResult<String>> results = new ArrayList<>();
        CsvReader.<String>mapping(row -> row.get("Name").asString())
                .duplicateHeaderPolicy(io.github.dornol.excelkit.core.read.DuplicateHeaderPolicy.FIRST)
                .read(new ByteArrayInputStream(csv.getBytes(StandardCharsets.UTF_8)), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success());
        assertEquals("Alice", results.get(0).data(), "First occurrence of duplicate header should be used");
    }
}
