package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.ReadResult;
import io.github.dornol.excelkit.excel.*;
import io.github.dornol.excelkit.excel.read.ExcelReadException;
import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.DisplayName;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

// ================================================================
// 2. ExcelReader.password(String) — encrypted Excel read
// ================================================================

@DisplayName("ExcelReader.password(String)")
class PasswordReadTest {

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
                .read(new ByteArrayInputStream(encrypted), results::add);

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
                        .read(new ByteArrayInputStream(encrypted), r -> {})
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
                        .read(new ByteArrayInputStream(encrypted), r -> {})
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
                        .read(new ByteArrayInputStream(out.toByteArray()), r -> {})
        );
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
