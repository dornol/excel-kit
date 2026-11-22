package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

// ================================================================
// 5. No-validator constructors: ExcelReader(Supplier), CsvReader(Supplier)
// ================================================================

@DisplayName("No-validator constructors")
class NoValidatorConstructorTest {

    @Test
    @DisplayName("ExcelReader(Supplier) works same as passing null validator")
    void excelReader_noValidatorConstructor() throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<User>create()
                .column("Name", u -> u.name)
                .write(Stream.of(makeUser("Alice", 30)))
                .writeTo(out);
        byte[] data = out.toByteArray();

        // With null validator
        List<ReadResult<User>> nullValidatorResults = new ArrayList<>();
        new ExcelReader<>(User::new, null)
                .column((u, cell) -> u.name = cell.asString())
                .read(new ByteArrayInputStream(data), nullValidatorResults::add);

        // With no-validator constructor
        List<ReadResult<User>> noValidatorResults = new ArrayList<>();
        new ExcelReader<>(User::new)
                .column((u, cell) -> u.name = cell.asString())
                .read(new ByteArrayInputStream(data), noValidatorResults::add);

        assertEquals(nullValidatorResults.size(), noValidatorResults.size());
        assertEquals(nullValidatorResults.get(0).data().name, noValidatorResults.get(0).data().name);
        assertTrue(nullValidatorResults.get(0).success());
        assertTrue(noValidatorResults.get(0).success());
    }

    @Test
    @DisplayName("CsvReader(Supplier) works same as passing null validator")
    void csvReader_noValidatorConstructor() {
        String csv = "Name\nAlice\n";
        byte[] data = csv.getBytes(StandardCharsets.UTF_8);

        // With null validator
        List<ReadResult<User>> nullValidatorResults = new ArrayList<>();
        new CsvReader<>(User::new, null)
                .column((u, cell) -> u.name = cell.asString())
                .read(new ByteArrayInputStream(data), nullValidatorResults::add);

        // With no-validator constructor
        List<ReadResult<User>> noValidatorResults = new ArrayList<>();
        new CsvReader<>(User::new)
                .column((u, cell) -> u.name = cell.asString())
                .read(new ByteArrayInputStream(data), noValidatorResults::add);

        assertEquals(nullValidatorResults.size(), noValidatorResults.size());
        assertEquals(nullValidatorResults.get(0).data().name, noValidatorResults.get(0).data().name);
        assertTrue(nullValidatorResults.get(0).success());
        assertTrue(noValidatorResults.get(0).success());
    }

    // ================================================================
    // Helper data class
    // ================================================================

    static class User {
        String name;
        int age;
    }

    private static User makeUser(String name, int age) {
        User u = new User();
        u.name = name;
        u.age = age;
        return u;
    }
}
