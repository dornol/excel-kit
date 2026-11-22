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
// 4. ExcelReader.setter() / CsvReader.setter() static factories
// ================================================================

@DisplayName("ExcelReader.setter() and CsvReader.setter() static factories")
class ReaderSetterFactoryTest {

    @Test
    @DisplayName("ExcelReader.setter() works identically to constructor")
    void excelReader_setter_worksLikeConstructor() throws IOException {
        // Write test data
        var out = new ByteArrayOutputStream();
        ExcelWriter.<User>create()
                .column("Name", u -> u.name)
                .column("Age", u -> u.age)
                .write(Stream.of(makeUser("Alice", 30)))
                .writeTo(out);
        byte[] data = out.toByteArray();

        // Read with constructor
        List<ReadResult<User>> constructorResults = new ArrayList<>();
        new ExcelReader<>(User::new)
                .column((u, cell) -> u.name = cell.asString())
                .column((u, cell) -> u.age = cell.asInt())
                .read(new ByteArrayInputStream(data), constructorResults::add);

        // Read with setter factory
        List<ReadResult<User>> setterResults = new ArrayList<>();
        ExcelReader.setter(User::new)
                .column((u, cell) -> u.name = cell.asString())
                .column((u, cell) -> u.age = cell.asInt())
                .read(new ByteArrayInputStream(data), setterResults::add);

        assertEquals(constructorResults.size(), setterResults.size());
        assertEquals(constructorResults.get(0).data().name, setterResults.get(0).data().name);
        assertEquals(constructorResults.get(0).data().age, setterResults.get(0).data().age);
    }

    @Test
    @DisplayName("CsvReader.setter() works identically to constructor")
    void csvReader_setter_worksLikeConstructor() {
        String csv = "Name,Age\nAlice,30\n";
        byte[] data = csv.getBytes(StandardCharsets.UTF_8);

        // Read with constructor
        List<ReadResult<User>> constructorResults = new ArrayList<>();
        new CsvReader<>(User::new)
                .column((u, cell) -> u.name = cell.asString())
                .column((u, cell) -> u.age = cell.asInt())
                .read(new ByteArrayInputStream(data), constructorResults::add);

        // Read with setter factory
        List<ReadResult<User>> setterResults = new ArrayList<>();
        CsvReader.setter(User::new)
                .column((u, cell) -> u.name = cell.asString())
                .column((u, cell) -> u.age = cell.asInt())
                .read(new ByteArrayInputStream(data), setterResults::add);

        assertEquals(constructorResults.size(), setterResults.size());
        assertEquals(constructorResults.get(0).data().name, setterResults.get(0).data().name);
        assertEquals(constructorResults.get(0).data().age, setterResults.get(0).data().age);
    }

    @Test
    @DisplayName("ExcelReader.setter() supports name-based columns")
    void excelReader_setter_withNamedColumns() throws IOException {
        var out = new ByteArrayOutputStream();
        ExcelWriter.<User>create()
                .column("Name", u -> u.name)
                .column("Age", u -> u.age)
                .write(Stream.of(makeUser("Bob", 42)))
                .writeTo(out);

        List<ReadResult<User>> results = new ArrayList<>();
        ExcelReader.setter(User::new)
                .column("Name", (u, cell) -> u.name = cell.asString())
                .column("Age", (u, cell) -> u.age = cell.asInt())
                .read(new ByteArrayInputStream(out.toByteArray()), results::add);

        assertEquals(1, results.size());
        assertEquals("Bob", results.get(0).data().name);
        assertEquals(42, results.get(0).data().age);
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
