package io.github.dornol.excelkit.core;

import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.csv.CsvWriteException;
import io.github.dornol.excelkit.excel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Tests for new features added in v0.14.0:
 * <ul>
 *   <li>{@link FileHandler#toFile(Path)} — default method on the FileHandler interface</li>
 *   <li>{@link ExcelReader#password(String)} — reading password-encrypted Excel files</li>
 *   <li>{@link CsvWriter#constColumnIf(String, boolean, Object)} — conditional constant column</li>
 *   <li>{@link ExcelReader#setter(java.util.function.Supplier)} / {@link CsvReader#setter(java.util.function.Supplier)} — static factory methods</li>
 *   <li>{@link ExcelReader(java.util.function.Supplier)} / {@link CsvReader(java.util.function.Supplier)} — no-validator constructors</li>
 * </ul>
 */
class NewFeaturesV014Test {

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

    // ================================================================
    // 1. FileHandler.writeTo(Path)
    // ================================================================

    @Nested
    @DisplayName("FileHandler.writeTo(Path)")
    class ToFileTests {

        @Test
        @DisplayName("ExcelHandler.toFile writes a valid .xlsx file")
        void excelHandler_toFile_writesValidXlsx(@TempDir Path tempDir) throws IOException {
            Path target = tempDir.resolve("output.xlsx");

            ExcelWriter.<String>create()
                    .column("Name", s -> s)
                    .write(Stream.of("Alice", "Bob"))
                    .writeTo(target);

            assertTrue(Files.exists(target));
            assertTrue(Files.size(target) > 0);
            try (var wb = new XSSFWorkbook(Files.newInputStream(target))) {
                assertEquals("Name", wb.getSheetAt(0).getRow(0).getCell(0).getStringCellValue());
                assertEquals("Alice", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                assertEquals("Bob", wb.getSheetAt(0).getRow(2).getCell(0).getStringCellValue());
            }
        }

        @Test
        @DisplayName("CsvHandler.toFile writes a valid .csv file")
        void csvHandler_toFile_writesValidCsv(@TempDir Path tempDir) throws IOException {
            Path target = tempDir.resolve("output.csv");

            new CsvWriter<String>()
                    .column("Name", s -> s)
                    .write(Stream.of("Alice", "Bob"))
                    .writeTo(target);

            assertTrue(Files.exists(target));
            String content = Files.readString(target, StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = content.split("\r?\n");
            assertEquals("Name", lines[0]);
            assertEquals("Alice", lines[1]);
            assertEquals("Bob", lines[2]);
        }

        @Test
        @DisplayName("Calling toFile() twice on ExcelHandler throws (one-shot contract)")
        void excelHandler_toFile_twice_throws(@TempDir Path tempDir) throws IOException {
            FileHandler handler = ExcelWriter.<String>create()
                    .column("Name", s -> s)
                    .write(Stream.of("Alice"));

            handler.writeTo(tempDir.resolve("first.xlsx"));

            assertThrows(ExcelWriteException.class,
                    () -> handler.writeTo(tempDir.resolve("second.xlsx")));
        }

        @Test
        @DisplayName("Calling toFile() twice on CsvHandler throws (one-shot contract)")
        void csvHandler_toFile_twice_throws(@TempDir Path tempDir) throws IOException {
            FileHandler handler = new CsvWriter<String>()
                    .column("Name", s -> s)
                    .write(Stream.of("Alice"));

            handler.writeTo(tempDir.resolve("first.csv"));

            assertThrows(CsvWriteException.class,
                    () -> handler.writeTo(tempDir.resolve("second.csv")));
        }

        @Test
        @DisplayName("toFile content matches write(OutputStream) content for Excel")
        void excelHandler_toFile_matchesWriteOutput(@TempDir Path tempDir) throws IOException {
            // Write via OutputStream
            var out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Val", s -> s)
                    .write(Stream.of("X"))
                    .writeTo(out);

            // Write via toFile
            Path target = tempDir.resolve("out.xlsx");
            ExcelWriter.<String>create()
                    .column("Val", s -> s)
                    .write(Stream.of("X"))
                    .writeTo(target);

            // Both should produce valid xlsx with identical data
            try (var wb1 = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()));
                 var wb2 = new XSSFWorkbook(Files.newInputStream(target))) {
                assertEquals(
                        wb1.getSheetAt(0).getRow(1).getCell(0).getStringCellValue(),
                        wb2.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            }
        }

        @Test
        @DisplayName("toFile content matches write(OutputStream) content for CSV")
        void csvHandler_toFile_matcheWriteOutput(@TempDir Path tempDir) throws IOException {
            // Write via OutputStream
            var out = new ByteArrayOutputStream();
            new CsvWriter<String>()
                    .column("Val", s -> s)
                    .write(Stream.of("X"))
                    .writeTo(out);

            // Write via toFile
            Path target = tempDir.resolve("out.csv");
            new CsvWriter<String>()
                    .column("Val", s -> s)
                    .write(Stream.of("X"))
                    .writeTo(target);

            String fromStream = out.toString(StandardCharsets.UTF_8);
            String fromFile = Files.readString(target, StandardCharsets.UTF_8);
            assertEquals(fromStream, fromFile);
        }

        @Test
        @DisplayName("Calling write() after toFile() throws (one-shot)")
        void writeAfterToFile_throws(@TempDir Path tempDir) throws IOException {
            FileHandler handler = ExcelWriter.<String>create()
                    .column("A", s -> s)
                    .write(Stream.of("x"));

            handler.writeTo(tempDir.resolve("out.xlsx"));

            assertThrows(ExcelWriteException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream()));
        }

        @Test
        @DisplayName("Calling toFile() after write() throws (one-shot)")
        void toFileAfterWrite_throws(@TempDir Path tempDir) throws IOException {
            FileHandler handler = ExcelWriter.<String>create()
                    .column("A", s -> s)
                    .write(Stream.of("x"));

            handler.writeTo(new ByteArrayOutputStream());

            assertThrows(ExcelWriteException.class,
                    () -> handler.writeTo(tempDir.resolve("out.xlsx")));
        }
    }

    // ================================================================
    // 2. ExcelReader.password(String) — encrypted Excel read
    // ================================================================

    @Nested
    @DisplayName("ExcelReader.password(String)")
    class PasswordReadTests {

        private byte[] writeEncryptedExcel(String password) throws IOException {
            var out = new ByteArrayOutputStream();
            ExcelWriter.<User>create()
                    .password(password)
                    .column("Name", u -> u.name)
                    .column("Age", u -> u.age)
                    .write(Stream.of(makeUser("Alice", 30), makeUser("Bob", 25)))
                    .writeTo(out);
            return out.toByteArray();
        }

        @Test
        @DisplayName("Write encrypted, read back with correct password")
        void readEncryptedWithCorrectPassword() throws IOException {
            byte[] encrypted = writeEncryptedExcel("secret");

            List<ReadResult<User>> results = new ArrayList<>();
            ExcelReader.setter(User::new)
                    .password("secret")
                    .column((u, cell) -> u.name = cell.asString())
                    .column((u, cell) -> u.age = cell.asInt())
                    .build(new ByteArrayInputStream(encrypted))
                    .read(results::add);

            assertEquals(2, results.size());
            assertEquals("Alice", results.get(0).data().name);
            assertEquals(30, results.get(0).data().age);
            assertEquals("Bob", results.get(1).data().name);
            assertEquals(25, results.get(1).data().age);
        }

        @Test
        @DisplayName("Wrong password throws ExcelReadException")
        void wrongPassword_throws() throws IOException {
            byte[] encrypted = writeEncryptedExcel("secret");

            ExcelReadException ex = assertThrows(ExcelReadException.class, () ->
                    ExcelReader.setter(User::new)
                            .password("wrong")
                            .column((u, cell) -> u.name = cell.asString())
                            .build(new ByteArrayInputStream(encrypted))
                            .read(r -> {})
            );
            assertTrue(ex.getMessage().contains("Invalid password"),
                    "Expected 'Invalid password' in message but got: " + ex.getMessage());
        }

        @Test
        @DisplayName("Reading encrypted file without password throws")
        void noPassword_throws() throws IOException {
            byte[] encrypted = writeEncryptedExcel("secret");

            // Without password, POI can't parse the POIFS file as OPC
            assertThrows(ExcelReadException.class, () ->
                    ExcelReader.setter(User::new)
                            .column((u, cell) -> u.name = cell.asString())
                            .build(new ByteArrayInputStream(encrypted))
                            .read(r -> {})
            );
        }

        @Test
        @DisplayName("Reading non-encrypted file with password throws")
        void passwordOnNonEncryptedFile_throws() throws IOException {
            // Write a plain (non-encrypted) Excel file
            var out = new ByteArrayOutputStream();
            ExcelWriter.<String>create()
                    .column("Val", s -> s)
                    .write(Stream.of("hello"))
                    .writeTo(out);

            // Reading a non-encrypted xlsx with password should fail
            // because the file is not a POIFS file
            assertThrows(ExcelReadException.class, () ->
                    ExcelReader.setter(User::new)
                            .password("secret")
                            .column((u, cell) -> u.name = cell.asString())
                            .build(new ByteArrayInputStream(out.toByteArray()))
                            .read(r -> {})
            );
        }
    }

    // ================================================================
    // 3. CsvWriter.constColumnIf
    // ================================================================

    @Nested
    @DisplayName("CsvWriter.constColumnIf")
    class CsvConstColumnIfTests {

        @Test
        @DisplayName("condition=true adds the constant column")
        void conditionTrue_addsColumn() throws IOException {
            var out = new ByteArrayOutputStream();
            new CsvWriter<String>()
                    .column("Name", s -> s)
                    .constColumnIf("Type", true, "USER")
                    .write(Stream.of("Alice"))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            assertEquals("Name,Type", lines[0]);
            assertEquals("Alice,USER", lines[1]);
        }

        @Test
        @DisplayName("condition=false skips the constant column")
        void conditionFalse_skipsColumn() throws IOException {
            var out = new ByteArrayOutputStream();
            new CsvWriter<String>()
                    .column("Name", s -> s)
                    .constColumnIf("Type", false, "USER")
                    .write(Stream.of("Alice"))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            assertEquals("Name", lines[0]);
            assertEquals("Alice", lines[1]);
        }

        @Test
        @DisplayName("constColumnIf with null value produces empty cell")
        void conditionTrue_nullValue() throws IOException {
            var out = new ByteArrayOutputStream();
            new CsvWriter<String>()
                    .column("Name", s -> s)
                    .constColumnIf("Extra", true, null)
                    .write(Stream.of("Alice"))
                    .writeTo(out);

            String csv = out.toString(StandardCharsets.UTF_8).replace("\uFEFF", "");
            String[] lines = csv.split("\r?\n");
            assertEquals("Name,Extra", lines[0]);
            assertEquals("Alice,", lines[1]);
        }
    }

    // ================================================================
    // 4. ExcelReader.setter() / CsvReader.setter() static factories
    // ================================================================

    @Nested
    @DisplayName("ExcelReader.setter() and CsvReader.setter() static factories")
    class SetterFactoryTests {

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
                    .build(new ByteArrayInputStream(data))
                    .read(constructorResults::add);

            // Read with setter factory
            List<ReadResult<User>> setterResults = new ArrayList<>();
            ExcelReader.setter(User::new)
                    .column((u, cell) -> u.name = cell.asString())
                    .column((u, cell) -> u.age = cell.asInt())
                    .build(new ByteArrayInputStream(data))
                    .read(setterResults::add);

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
                    .build(new ByteArrayInputStream(data))
                    .read(constructorResults::add);

            // Read with setter factory
            List<ReadResult<User>> setterResults = new ArrayList<>();
            CsvReader.setter(User::new)
                    .column((u, cell) -> u.name = cell.asString())
                    .column((u, cell) -> u.age = cell.asInt())
                    .build(new ByteArrayInputStream(data))
                    .read(setterResults::add);

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
                    .build(new ByteArrayInputStream(out.toByteArray()))
                    .read(results::add);

            assertEquals(1, results.size());
            assertEquals("Bob", results.get(0).data().name);
            assertEquals(42, results.get(0).data().age);
        }
    }

    // ================================================================
    // 5. No-validator constructors: ExcelReader(Supplier), CsvReader(Supplier)
    // ================================================================

    @Nested
    @DisplayName("No-validator constructors")
    class NoValidatorConstructorTests {

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
                    .build(new ByteArrayInputStream(data))
                    .read(nullValidatorResults::add);

            // With no-validator constructor
            List<ReadResult<User>> noValidatorResults = new ArrayList<>();
            new ExcelReader<>(User::new)
                    .column((u, cell) -> u.name = cell.asString())
                    .build(new ByteArrayInputStream(data))
                    .read(noValidatorResults::add);

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
                    .build(new ByteArrayInputStream(data))
                    .read(nullValidatorResults::add);

            // With no-validator constructor
            List<ReadResult<User>> noValidatorResults = new ArrayList<>();
            new CsvReader<>(User::new)
                    .column((u, cell) -> u.name = cell.asString())
                    .build(new ByteArrayInputStream(data))
                    .read(noValidatorResults::add);

            assertEquals(nullValidatorResults.size(), noValidatorResults.size());
            assertEquals(nullValidatorResults.get(0).data().name, noValidatorResults.get(0).data().name);
            assertTrue(nullValidatorResults.get(0).success());
            assertTrue(noValidatorResults.get(0).success());
        }
    }
}
