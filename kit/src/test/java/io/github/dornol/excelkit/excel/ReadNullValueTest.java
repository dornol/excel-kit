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

@DisplayName("B1: Writer nullValue")
class ReadNullValueTest {

    @Test
    void nullValue_throughColumnBuilder_shouldWork() throws IOException {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        writer.column("Name", s -> s, c -> c.nullValue("-"));

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        writer.write(Stream.of((String) null)).writeTo(bos);

        // Read back to verify
        List<ReadResult<String>> results = new ArrayList<>();
        ExcelReader.<String>mapping(row -> row.get("Name").asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertTrue(results.get(0).success());
        assertEquals("-", results.get(0).data());
    }

    @Test
    void nullValue_throughDefaultStyle_shouldFallback() throws IOException {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        writer.defaultStyle(d -> d.nullValue("DEFAULT"));
        writer.column("Col", s -> s);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        writer.write(Stream.of((String) null)).writeTo(bos);

        List<ReadResult<String>> results = new ArrayList<>();
        ExcelReader.<String>mapping(row -> row.get("Col").asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertEquals("DEFAULT", results.get(0).data());
    }

    @Test
    void nullValue_perColumn_overridesDefault() throws IOException {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        writer.defaultStyle(d -> d.nullValue("DEFAULT"));
        writer.column("Col", s -> s, c -> c.nullValue("CUSTOM"));

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        writer.write(Stream.of((String) null)).writeTo(bos);

        List<ReadResult<String>> results = new ArrayList<>();
        ExcelReader.<String>mapping(row -> row.get("Col").asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertEquals("CUSTOM", results.get(0).data());
    }

    @Test
    void nullValue_nonNullData_shouldUseActualData() throws IOException {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        writer.column("Name", s -> s, c -> c.nullValue("N/A"));

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        writer.write(Stream.of("Hello")).writeTo(bos);

        List<ReadResult<String>> results = new ArrayList<>();
        ExcelReader.<String>mapping(row -> row.get("Name").asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals("Hello", results.get(0).data());
    }

    @Test
    void nullValue_withoutSetting_shouldWriteEmpty() throws IOException {
        ExcelWriter<String> writer = ExcelWriter.<String>create();
        writer.column("Name", s -> s);

        ByteArrayOutputStream bos = new ByteArrayOutputStream();
        writer.write(Stream.of((String) null)).writeTo(bos);

        List<ReadResult<String>> results = new ArrayList<>();
        ExcelReader.<String>mapping(row -> row.get("Name").asString())
                .read(new ByteArrayInputStream(bos.toByteArray()), results::add);

        assertEquals("", results.get(0).data());
    }
}
