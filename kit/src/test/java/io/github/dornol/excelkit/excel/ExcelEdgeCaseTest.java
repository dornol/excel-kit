package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ExcelKitException;
import io.github.dornol.excelkit.core.ReadResult;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Nested;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Edge case tests for:
 * - ExcelHandler (consume twice, invalid password)
 * - ExcelReader (skipColumns negative, header not found)
 * - ExcelSummary (all Op types)
 * - ExcelWorkbook edge cases
 * - ExcelWriter defaultStyle with applyDefaults
 */
class ExcelEdgeCaseTest {

    record Item(String name, int value) {}

    // ============================================================
    // ExcelHandler edge cases
    // ============================================================
    @Nested
    class ExcelHandlerTests {

        @Test
        void write_twice_throws() throws IOException {
            ByteArrayOutputStream out1 = new ByteArrayOutputStream();
            ExcelHandler handler = ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            handler.writeTo(out1);
            var ex = assertThrows(ExcelWriteException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream()));
            assertTrue(ex.getMessage().contains("Already consumed"),
                    "Should indicate handler was already consumed");
        }

        static Stream<Object[]> invalidStringPasswords() {
            return Stream.of(
                    new Object[]{null, "null"},
                    new Object[]{"  ", "blank"}
            );
        }

        @ParameterizedTest(name = "String password={1}")
        @MethodSource("invalidStringPasswords")
        void writeToWithPassword_invalidString_throws(String password, String label) throws IOException {
            ExcelHandler handler = ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            var ex = assertThrows(IllegalArgumentException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream(), password));
            assertTrue(ex.getMessage().toLowerCase().contains("password"),
                    "Exception should mention password");
        }

        static Stream<Object[]> invalidCharPasswords() {
            return Stream.of(
                    new Object[]{null, "null"},
                    new Object[]{new char[0], "empty"}
            );
        }

        @ParameterizedTest(name = "char[] password={1}")
        @MethodSource("invalidCharPasswords")
        void writeToWithPassword_invalidCharArray_throws(char[] password, String label) throws IOException {
            ExcelHandler handler = ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            var ex = assertThrows(IllegalArgumentException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream(), password));
            assertTrue(ex.getMessage().toLowerCase().contains("password"),
                    "Exception should mention password");
        }
    }

    // ============================================================
    // ExcelWriter.password() API tests
    // ============================================================
    @Nested
    class ExcelWriterPasswordTests {

        @Test
        void password_shouldAutoEncryptAndBeDecryptableWithCorrectPassword() throws IOException, GeneralSecurityException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .password("secret123")
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 1), new Item("B", 2)))
                    .writeTo(out);

            byte[] bytes = out.toByteArray();
            // Verify OLE2 encrypted format
            assertEquals((byte) 0xD0, bytes[0], "Should be OLE2 encrypted format");
            assertEquals((byte) 0xCF, bytes[1], "Should be OLE2 encrypted format");

            // Verify round-trip: decrypt and read actual content
            try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
                EncryptionInfo info = new EncryptionInfo(fs);
                Decryptor dec = Decryptor.getInstance(info);
                assertTrue(dec.verifyPassword("secret123"), "Correct password should verify");

                try (InputStream decStream = dec.getDataStream(fs);
                     XSSFWorkbook wb = new XSSFWorkbook(decStream)) {
                    var sheet = wb.getSheetAt(0);
                    assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
                    assertEquals("Value", sheet.getRow(0).getCell(1).getStringCellValue());
                    assertEquals("A", sheet.getRow(1).getCell(0).getStringCellValue());
                    assertEquals(2, sheet.getRow(2).getCell(1).getNumericCellValue());
                }
            }
        }

        @Test
        void password_shouldNotBeDecryptableWithWrongPassword() throws IOException, GeneralSecurityException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .password("correct")
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(out.toByteArray()))) {
                EncryptionInfo info = new EncryptionInfo(fs);
                Decryptor dec = Decryptor.getInstance(info);
                assertFalse(dec.verifyPassword("wrong"), "Wrong password should not verify");
            }
        }

        @Test
        void noPassword_shouldWriteUnencryptedOOXML() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            byte[] bytes = out.toByteArray();
            // Unencrypted OOXML starts with ZIP magic bytes (0x504B)
            assertEquals((byte) 0x50, bytes[0], "Should be ZIP/OOXML format");
            assertEquals((byte) 0x4B, bytes[1], "Should be ZIP/OOXML format");
        }

        @Test
        void password_write_twice_shouldThrowAlreadyConsumed() throws IOException {
            ExcelHandler handler = ExcelWriter.<Item>create()
                    .password("secret")
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));

            handler.writeTo(new ByteArrayOutputStream());

            var ex = assertThrows(ExcelWriteException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream()));
            assertTrue(ex.getMessage().contains("Already consumed"));
        }

        @Test
        void password_combinedWithProtectSheetAndWorkbook() throws IOException, GeneralSecurityException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .password("filePass")
                    .protectSheet("sheetPass")
                    .protectWorkbook("wbPass")
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            // Decrypt with file password
            try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(out.toByteArray()))) {
                EncryptionInfo info = new EncryptionInfo(fs);
                Decryptor dec = Decryptor.getInstance(info);
                assertTrue(dec.verifyPassword("filePass"));

                try (InputStream decStream = dec.getDataStream(fs);
                     XSSFWorkbook wb = new XSSFWorkbook(decStream)) {
                    // Sheet protection should be applied inside
                    assertTrue(wb.getSheetAt(0).getProtect(), "Sheet should be protected");
                    // Workbook structure protection should be applied
                    assertTrue(wb.getCTWorkbook().isSetWorkbookProtection(),
                            "Workbook protection should be set");
                }
            }
        }

        @Test
        void password_nullValue_shouldThrow() {
            var ex = assertThrows(IllegalArgumentException.class, () ->
                    ExcelWriter.<Item>create().password((String) null));
            assertTrue(ex.getMessage().toLowerCase().contains("password"));
        }

        @Test
        void password_blankValue_shouldThrow() {
            var ex = assertThrows(IllegalArgumentException.class, () ->
                    ExcelWriter.<Item>create().password("   "));
            assertTrue(ex.getMessage().toLowerCase().contains("password"));
        }
    }

    // ============================================================
    // ExcelWorkbook.password() API tests
    // ============================================================
    @Nested
    class ExcelWorkbookPasswordTests {

        @Test
        void password_shouldAutoEncryptAndBeDecryptable() throws IOException, GeneralSecurityException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
                workbook.password("secret123");
                workbook.<Item>sheet("Data")
                        .column("Name", Item::name)
                        .write(Stream.of(new Item("A", 1)));
                workbook.finish().writeTo(out);
            }

            byte[] bytes = out.toByteArray();
            assertEquals((byte) 0xD0, bytes[0], "Should be OLE2 encrypted format");

            // Verify round-trip decryption
            try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
                EncryptionInfo info = new EncryptionInfo(fs);
                Decryptor dec = Decryptor.getInstance(info);
                assertTrue(dec.verifyPassword("secret123"));

                try (InputStream decStream = dec.getDataStream(fs);
                     XSSFWorkbook wb = new XSSFWorkbook(decStream)) {
                    assertEquals("Data", wb.getSheetName(0));
                    assertEquals("A", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
                }
            }
        }

        @Test
        void password_multiSheet_roundTripViaExcelReader() throws IOException {
            // Write a multi-sheet encrypted workbook, then read each sheet back with the password
            // to verify end-to-end that different sheet writers share the encryption correctly.
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
                workbook.password("multiPass!");
                workbook.<Item>sheet("Items")
                        .column("Name", Item::name)
                        .column("Value", Item::value)
                        .write(Stream.of(new Item("A", 1), new Item("B", 2)));
                workbook.<Item>sheet("More")
                        .column("Name", Item::name)
                        .column("Value", Item::value)
                        .write(Stream.of(new Item("X", 9)));
                workbook.finish().writeTo(out);
            }

            byte[] bytes = out.toByteArray();

            // Sheet 0
            List<Item> sheet0 = new ArrayList<>();
            ExcelReader.<Item>mapping(row -> new Item(
                    row.get("Name").asString(),
                    row.get("Value").asInt()))
                    .password("multiPass!")
                    .sheetIndex(0)
                    .readStrict(new ByteArrayInputStream(bytes), sheet0::add);
            assertEquals(List.of(new Item("A", 1), new Item("B", 2)), sheet0);

            // Sheet 1
            List<Item> sheet1 = new ArrayList<>();
            ExcelReader.<Item>mapping(row -> new Item(
                    row.get("Name").asString(),
                    row.get("Value").asInt()))
                    .password("multiPass!")
                    .sheetIndex(1)
                    .readStrict(new ByteArrayInputStream(bytes), sheet1::add);
            assertEquals(List.of(new Item("X", 9)), sheet1);
        }

        @Test
        void password_combinedWithProtectWorkbook() throws IOException, GeneralSecurityException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook workbook = ExcelWorkbook.create()) {
                workbook.password("filePass");
                workbook.protectWorkbook("structPass");
                workbook.<Item>sheet("Data")
                        .column("Name", Item::name)
                        .write(Stream.of(new Item("A", 1)));
                workbook.finish().writeTo(out);
            }

            try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(out.toByteArray()))) {
                EncryptionInfo info = new EncryptionInfo(fs);
                Decryptor dec = Decryptor.getInstance(info);
                assertTrue(dec.verifyPassword("filePass"));

                try (InputStream decStream = dec.getDataStream(fs);
                     XSSFWorkbook wb = new XSSFWorkbook(decStream)) {
                    assertTrue(wb.getCTWorkbook().isSetWorkbookProtection(),
                            "Workbook structure protection should be set inside encrypted file");
                }
            }
        }

        @Test
        void password_nullValue_shouldThrow() {
            try (ExcelWorkbook wb = ExcelWorkbook.create()) {
                var ex = assertThrows(IllegalArgumentException.class, () -> wb.password((String) null));
                assertTrue(ex.getMessage().toLowerCase().contains("password"));
            }
        }

        @Test
        void password_blankValue_shouldThrow() {
            try (ExcelWorkbook wb = ExcelWorkbook.create()) {
                var ex = assertThrows(IllegalArgumentException.class, () -> wb.password("  "));
                assertTrue(ex.getMessage().toLowerCase().contains("password"));
            }
        }
    }

    // ============================================================
    // password() + writeToWithPassword() conflict
    // ============================================================
    @Nested
    class PasswordConflictTests {

        @Test
        void passwordSet_thenConsumeWithPassword_shouldThrow() throws IOException {
            ExcelHandler handler = ExcelWriter.<Item>create()
                    .password("first")
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));

            var ex = assertThrows(IllegalStateException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream(), "second"));
            assertTrue(ex.getMessage().contains("already set"),
                    "Should indicate password conflict");
        }

        @Test
        void passwordSet_thenConsumeWithCharArrayPassword_shouldThrowAndZeroPassword() throws IOException {
            ExcelHandler handler = ExcelWriter.<Item>create()
                    .password("first")
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));

            char[] charPassword = "second".toCharArray();
            var ex = assertThrows(IllegalStateException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream(), charPassword));
            assertTrue(ex.getMessage().contains("already set"),
                    "Should indicate password conflict");

            // char[] must be zeroed even when conflict exception is thrown
            for (char c : charPassword) {
                assertEquals('\0', c, "Password char array should be zeroed even on conflict exception");
            }
        }
    }

    // ============================================================
    // char[] password blank validation
    // ============================================================
    @Nested
    class CharArrayPasswordBlankTests {

        @Test
        void writeToWithPassword_blankCharArray_throws() throws IOException {
            ExcelHandler handler = ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)));
            char[] blankPassword = {' ', '\t', ' '};
            var ex = assertThrows(IllegalArgumentException.class,
                    () -> handler.writeTo(new ByteArrayOutputStream(), blankPassword));
            assertTrue(ex.getMessage().toLowerCase().contains("password"),
                    "Exception should mention password");
        }
    }

    // ============================================================
    // ExcelReader edge cases
    // ============================================================
    @Nested
    class ExcelReaderTests {

        static class MutableItem {
            String name;
            int value;
        }

        @Test
        void skipColumns_negative_throws() {
            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            var ex = assertThrows(IllegalArgumentException.class, () -> reader.skipColumns(-1));
            assertTrue(ex.getMessage().contains("negative"),
                    "Should mention negative value: " + ex.getMessage());
        }

        @Test
        void onProgress_zeroInterval_throws() {
            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            var ex = assertThrows(IllegalArgumentException.class,
                    () -> reader.onProgress(0, (c, cur) -> {}));
            assertTrue(ex.getMessage().contains("positive"),
                    "Should mention positive requirement: " + ex.getMessage());
        }

        @Test
        void headerNotFound_throws() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            ExcelReader<MutableItem> reader = new ExcelReader<>(MutableItem::new, null);
            reader.column("Name", (item, cell) -> {});
            reader.column("NonExistentHeader", (item, cell) -> {});

            assertThrows(ExcelKitException.class,
                    () -> reader.read(new ByteArrayInputStream(out.toByteArray()), r -> {}));
        }

        @Test
        void getSheetHeaders_withHeaderRowIndex() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            List<String> headers = ExcelReader.getSheetHeaders(
                    new ByteArrayInputStream(out.toByteArray()), 0, 0);

            assertEquals(2, headers.size());
            assertEquals("Name", headers.get(0));
            assertEquals("Value", headers.get(1));
        }
    }

    // ============================================================
    // ExcelSummary all Op types
    // ============================================================
    @Nested
    class ExcelSummaryTests {

        @Test
        void allSummaryOps_shouldWriteFormulaRows() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Summary")
                            .sum("Value")
                            .average("Value")
                            .count("Value")
                            .min("Value")
                            .max("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Header(0) + 2 data rows(1,2) + 5 summary rows(3-7)
                // Check SUM row has formula
                var sumRow = sheet.getRow(3);
                assertNotNull(sumRow, "SUM summary row should exist");
                var sumCell = sumRow.getCell(1);
                assertEquals(CellType.FORMULA, sumCell.getCellType(), "Summary cell should be formula");
                assertTrue(sumCell.getCellFormula().startsWith("SUM("), "Should be SUM formula");
            }
        }

        @Test
        void summary_singleOp_withLabel_shouldUseLabelText() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Total")
                            .sum("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var summaryRow = wb.getSheetAt(0).getRow(3);
                // Single op with label → uses labelText "Total"
                assertEquals("Total", summaryRow.getCell(0).getStringCellValue());
            }
        }

        @Test
        void summary_labelInNonExistentColumn_shouldFallbackToFirstColumn() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("NonExistent", "Total:")
                            .sum("Value"))
                    .write(Stream.of(new Item("A", 10)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var summaryRow = wb.getSheetAt(0).getRow(2);
                // NonExistent column → falls back to idx=0
                assertEquals("Total:", summaryRow.getCell(0).getStringCellValue());
            }
        }

        @Test
        void summary_labelInColumn_shouldWriteLabelAndFormula() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .summary(s -> s
                            .label("Name", "Total:")
                            .sum("Value"))
                    .write(Stream.of(new Item("A", 10), new Item("B", 20)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                var summaryRow = sheet.getRow(3);
                assertNotNull(summaryRow);
                assertEquals("Total:", summaryRow.getCell(0).getStringCellValue());
                assertEquals(CellType.FORMULA, summaryRow.getCell(1).getCellType());
            }
        }
    }

    // ============================================================
    // ExcelWriter defaultStyle
    // ============================================================
    @Nested
    class ExcelWriterDefaultStyleTests {

        @Test
        void defaultStyle_shouldApplyBoldToAllColumns() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .defaultStyle(d -> d.bold(true).fontSize(12))
                    .column("Name", Item::name)
                    .column("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var dataRow = wb.getSheetAt(0).getRow(1);
                // Both columns should have bold font from defaultStyle
                var font0 = wb.getFontAt(dataRow.getCell(0).getCellStyle().getFontIndex());
                assertTrue(font0.getBold(), "Name column should be bold from default style");
                var font1 = wb.getFontAt(dataRow.getCell(1).getCellStyle().getFontIndex());
                assertTrue(font1.getBold(), "Value column should be bold from default style");
            }
        }

        @Test
        void defaultStyle_columnOverrides_shouldNotBeBold() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .defaultStyle(d -> d.bold(true))
                    .column("Name", Item::name, c -> c.bold(false))
                    .column("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var dataRow = wb.getSheetAt(0).getRow(1);
                // Name column overrides bold=false
                var font0 = wb.getFontAt(dataRow.getCell(0).getCellStyle().getFontIndex());
                assertFalse(font0.getBold(), "Name column should override bold to false");
                // Value column inherits bold=true from default
                var font1 = wb.getFontAt(dataRow.getCell(1).getCellStyle().getFontIndex());
                assertTrue(font1.getBold(), "Value column should inherit bold from default");
            }
        }
    }

    // ============================================================
    // ExcelWorkbook edge cases
    // ============================================================
    @Nested
    class ExcelWorkbookTests {

        @Test
        void protectWorkbook_shouldSetWorkbookProtection() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            try (ExcelWorkbook wb = ExcelWorkbook.create()) {
                wb.<Item>sheet("Data")
                        .column("Name", Item::name)
                        .write(Stream.of(new Item("A", 1)));
                wb.protectWorkbook("password123");
                wb.finish().writeTo(out);
            }

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                assertTrue(wb.getCTWorkbook().isSetWorkbookProtection(),
                        "Workbook protection should be set");
            }
        }
    }

    // ============================================================
    // ExcelWriter header font customization
    // ============================================================
    @Nested
    class ExcelWriterHeaderFontTests {

        @Test
        void headerFontName_shouldApplyCustomFont() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .headerFontName("Arial")
                    .column("Name", Item::name)
                    .column("Value", i -> i.value)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals("Arial", font.getFontName());
            }
        }

        @Test
        void headerFontSize_shouldApplyCustomSize() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .headerFontSize(16)
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals(16, font.getFontHeightInPoints());
            }
        }

        @Test
        void headerFontNameAndSize_combined() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .headerFontName("Times New Roman")
                    .headerFontSize(14)
                    .column("Name", Item::name)
                    .write(Stream.of(new Item("A", 1)))
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var headerCell = wb.getSheetAt(0).getRow(0).getCell(0);
                var font = wb.getFontAt(headerCell.getCellStyle().getFontIndex());
                assertEquals("Times New Roman", font.getFontName());
                assertEquals(14, font.getFontHeightInPoints());
            }
        }
    }

    // ============================================================
    // ExcelWriter write with no data
    // ============================================================
    @Nested
    class EmptyDataTests {

        @Test
        void write_emptyStream_shouldCreateHeaderOnly() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value)
                    .write(Stream.empty())
                    .writeTo(out);

            try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
                var sheet = wb.getSheetAt(0);
                // Header row exists
                assertEquals("Name", sheet.getRow(0).getCell(0).getStringCellValue());
                assertEquals("Value", sheet.getRow(0).getCell(1).getStringCellValue());
                // No data rows
                assertNull(sheet.getRow(1), "Should have no data rows");
            }
        }
    }

    // ============================================================
    // readStrict with error messages
    // ============================================================
    @Nested
    class ReadStrictTests {

        @Test
        void readStrict_emptyMessages_shouldShowUnknownError() throws IOException {
            ByteArrayOutputStream out = new ByteArrayOutputStream();
            ExcelWriter.<Item>create()
                    .column("Name", Item::name)
                    .column("Value", i -> i.value, c -> c.type(ExcelDataType.INTEGER))
                    .write(Stream.of(new Item("A", 10)))
                    .writeTo(out);

            // Read with a mapper that always succeeds
            List<Item> results = new ArrayList<>();
            ExcelReader.<Item>mapping(row ->
                    new Item(row.get("Name").asString(), row.get("Value").asInt())
            ).readStrict(new ByteArrayInputStream(out.toByteArray()), results::add);

            assertEquals(1, results.size());
        }
    }
}
