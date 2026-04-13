package io.github.dornol.excelkit.excel;

import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.SheetConditionalFormatting;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.security.GeneralSecurityException;
import java.util.List;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Integration tests that verify POI features survive write → read cycles
 * in combinations that isolated unit tests don't exercise:
 * <ul>
 *   <li>Encryption + dropdown validation + cell comment + conditional formatting
 *       all present in the same file (do any get silently dropped under encryption?)</li>
 *   <li>Conditional formatting + data validation coexisting on the same sheet
 *       (do they interfere at the XLSX level?)</li>
 * </ul>
 * These guard against regressions where a POI version change or internal
 * reordering causes one feature to silently vanish.
 */
class FeatureIntegrationRoundTripTest {

    record Person(String name, int age, String status) {}

    // ──────────────────────────────────────────────────────────────
    // 1. Encryption + rich features round-trip
    // ──────────────────────────────────────────────────────────────

    @Test
    void encrypted_withValidationCommentAndConditionalFormatting_roundTrip()
            throws IOException, GeneralSecurityException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Person>create()
                .password("featurePw")
                .column("Name", Person::name, c -> c.comment(p -> "Name: " + p.name()))
                .column("Age", Person::age, c -> c.type(ExcelDataType.INTEGER))
                .column("Status", Person::status, c -> c.dropdown("Active", "Inactive"))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("30", ExcelColor.LIGHT_RED))
                .write(Stream.of(
                        new Person("Alice", 28, "Active"),
                        new Person("Bob", 45, "Inactive")))
                .writeTo(out);

        byte[] bytes = out.toByteArray();
        // Verify the file is encrypted (OLE2 magic bytes)
        assertEquals((byte) 0xD0, bytes[0], "Expected encrypted OLE2 file");
        assertEquals((byte) 0xCF, bytes[1], "Expected encrypted OLE2 file");

        // Decrypt via POI and re-open as XSSFWorkbook to check all features
        try (POIFSFileSystem fs = new POIFSFileSystem(new ByteArrayInputStream(bytes))) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor dec = Decryptor.getInstance(info);
            assertTrue(dec.verifyPassword("featurePw"));

            try (InputStream decStream = dec.getDataStream(fs);
                 XSSFWorkbook wb = new XSSFWorkbook(decStream)) {
                XSSFSheet sheet = wb.getSheetAt(0);

                // Data preserved
                assertEquals("Alice", sheet.getRow(1).getCell(0).getStringCellValue());
                assertEquals(28, (int) sheet.getRow(1).getCell(1).getNumericCellValue());
                assertEquals("Bob", sheet.getRow(2).getCell(0).getStringCellValue());

                // Cell comment preserved
                var comment = sheet.getRow(1).getCell(0).getCellComment();
                assertNotNull(comment, "Cell comment should survive encryption round-trip");
                assertEquals("Name: Alice", comment.getString().getString());

                // Dropdown validation preserved
                List<? extends DataValidation> validations = sheet.getDataValidations();
                assertFalse(validations.isEmpty(), "Dropdown validation should survive encryption");

                // Conditional formatting preserved
                SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
                assertTrue(scf.getNumConditionalFormattings() > 0,
                        "Conditional formatting should survive encryption");
            }
        }
    }

    @Test
    void encrypted_dataReadableViaExcelReader_withRichFeatures() throws IOException {
        // Complementary check: the event-based ExcelReader (which ignores formatting)
        // must still extract cell values correctly from a rich encrypted file.
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Person>create()
                .password("pw")
                .column("Name", Person::name, c -> c.comment(p -> "c"))
                .column("Age", Person::age, c -> c.type(ExcelDataType.INTEGER))
                .column("Status", Person::status, c -> c.dropdown("A", "B"))
                .conditionalFormatting(cf -> cf.columns(1).greaterThan("30", ExcelColor.LIGHT_RED))
                .write(Stream.of(new Person("Alice", 28, "A"), new Person("Bob", 45, "B")))
                .writeTo(out);

        java.util.List<Person> read = new java.util.ArrayList<>();
        ExcelReader.<Person>mapping(row -> new Person(
                row.get("Name").asString(),
                row.get("Age").asInt(),
                row.get("Status").asString()))
                .password("pw")
                .build(new ByteArrayInputStream(out.toByteArray()))
                .readStrict(read::add);

        assertEquals(
                java.util.List.of(new Person("Alice", 28, "A"), new Person("Bob", 45, "B")),
                read);
    }

    // ──────────────────────────────────────────────────────────────
    // 2. Conditional formatting + data validation coexistence
    // ──────────────────────────────────────────────────────────────

    @Test
    void conditionalFormattingAndDataValidation_onSameSheet_bothPreserved() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Person>create()
                .column("Name", Person::name)
                .column("Age", Person::age, c -> c.type(ExcelDataType.INTEGER))
                .column("Status", Person::status, c -> c.dropdown("Active", "Inactive"))
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("40", ExcelColor.LIGHT_RED))
                .write(Stream.of(
                        new Person("Alice", 30, "Active"),
                        new Person("Bob", 50, "Inactive")))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFSheet sheet = wb.getSheetAt(0);

            List<? extends DataValidation> validations = sheet.getDataValidations();
            assertFalse(validations.isEmpty(),
                    "Data validation (dropdown) must not be dropped when conditional formatting is also set");

            SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
            assertTrue(scf.getNumConditionalFormattings() > 0,
                    "Conditional formatting must not be dropped when data validation is also set");

            // Spot-check validation applies to Status column (index 2)
            boolean dropdownCoversStatus = validations.stream()
                    .flatMap(v -> java.util.Arrays.stream(v.getRegions().getCellRangeAddresses()))
                    .anyMatch(r -> r.getFirstColumn() <= 2 && r.getLastColumn() >= 2);
            assertTrue(dropdownCoversStatus,
                    "Dropdown validation should target the Status column (index 2)");

            // Spot-check conditional formatting applies to Age column (index 1)
            var cf = scf.getConditionalFormattingAt(0);
            boolean condCoversAge = java.util.Arrays.stream(cf.getFormattingRanges())
                    .anyMatch(r -> r.getFirstColumn() <= 1 && r.getLastColumn() >= 1);
            assertTrue(condCoversAge,
                    "Conditional formatting should target the Age column (index 1)");
        }
    }

    @Test
    void conditionalFormattingAndDataValidation_onOverlappingColumn_bothApplied()
            throws IOException {
        // Both features targeting the SAME column — make sure POI doesn't drop one.
        // Conditional formatting on column 1 (Age), plus a numeric validation on
        // column 1 to constrain the allowed range.
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<Person>create()
                .column("Name", Person::name)
                .column("Age", Person::age, c -> c
                        .type(ExcelDataType.INTEGER)
                        .validation(ExcelValidation.integerBetween(0, 120)))
                .column("Status", Person::status)
                .conditionalFormatting(cf -> cf
                        .columns(1)
                        .greaterThan("65", ExcelColor.LIGHT_RED))
                .write(Stream.of(new Person("Alice", 30, "A"), new Person("Bob", 70, "B")))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            XSSFSheet sheet = wb.getSheetAt(0);

            // Both features present on the same column
            List<? extends DataValidation> validations = sheet.getDataValidations();
            assertFalse(validations.isEmpty(), "Numeric validation must survive");

            boolean validationOnAge = validations.stream()
                    .flatMap(v -> java.util.Arrays.stream(v.getRegions().getCellRangeAddresses()))
                    .anyMatch(r -> r.getFirstColumn() <= 1 && r.getLastColumn() >= 1);
            assertTrue(validationOnAge, "Numeric validation should cover Age column");

            SheetConditionalFormatting scf = sheet.getSheetConditionalFormatting();
            assertTrue(scf.getNumConditionalFormattings() > 0,
                    "Conditional formatting must survive");
        }
    }
}
