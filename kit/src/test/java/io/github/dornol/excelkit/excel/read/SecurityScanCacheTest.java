package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadSecurityException;
import io.github.dornol.excelkit.core.read.ReadSecurityPolicy;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

/** The security scan of a workbook file happens once per policy, whether it passed or refused. */
class SecurityScanCacheTest {

    @TempDir
    Path dir;

    private Path workbook(String name, boolean withFormula) throws IOException {
        Path file = dir.resolve(name);
        try (var wb = new XSSFWorkbook(); OutputStream out = Files.newOutputStream(file)) {
            var row = wb.createSheet("S").createRow(0);
            row.createCell(0).setCellValue("a");
            if (withFormula) row.createCell(1).setCellFormula("1+1");
            wb.write(out);
        }
        return file;
    }

    private void replace(Path file, Path content) throws IOException {
        Files.copy(content, file, java.nio.file.StandardCopyOption.REPLACE_EXISTING);
    }

    @Test
    void aCleanResultIsNotScannedAgainForTheSamePolicy() throws Exception {
        Path file = workbook("clean.xlsx", false);
        var scans = new SecurityScanCache();
        scans.scan(file, ReadSecurityPolicy.STRICT);

        // If the file were scanned again, the formula that is in it now would be refused.
        replace(file, workbook("formula.xlsx", true));

        assertDoesNotThrow(() -> scans.scan(file, ReadSecurityPolicy.STRICT));
        assertThrows(ReadSecurityException.class, () -> new SecurityScanCache().scan(file, ReadSecurityPolicy.STRICT));
    }

    @Test
    void aRefusalIsRememberedAndRethrown() throws Exception {
        Path file = workbook("formula.xlsx", true);
        var scans = new SecurityScanCache();
        var first = assertThrows(ReadSecurityException.class, () -> scans.scan(file, ReadSecurityPolicy.STRICT));
        assertEquals(ReadSecurityException.Reason.FORMULA, first.reason());

        replace(file, workbook("clean.xlsx", false));

        var second = assertThrows(ReadSecurityException.class, () -> scans.scan(file, ReadSecurityPolicy.STRICT));
        assertSame(first, second, "the same refusal, without another scan");
    }

    @Test
    void eachPolicyIsScannedOnItsOwn() throws Exception {
        Path file = workbook("formula.xlsx", true);
        var scans = new SecurityScanCache();

        assertThrows(ReadSecurityException.class, () -> scans.scan(file, ReadSecurityPolicy.STRICT));
        assertDoesNotThrow(() -> scans.scan(file, ReadSecurityPolicy.DEFAULT));
    }

    @Test
    void everySheetOfAWorkbookWithAFormulaIsRefusedUnderTheStrictPolicy() throws Exception {
        Path file = workbook("formula.xlsx", true);
        var reader = (java.util.function.Supplier<ExcelReader<java.util.Map<String, String>>>) () ->
                ExcelReader.forMap().securityPolicy(ReadSecurityPolicy.STRICT);

        WorkbookReadReport report = ExcelWorkbookReader.create()
                .sheet(0, reader.get(), result -> fail("no rows before the scan passes"))
                .sheet("S", reader.get(), result -> fail("no rows before the scan passes"))
                .read(file);

        assertEquals(2, report.sheets().size());
        for (SheetReadReport sheet : report.sheets()) {
            assertInstanceOf(ReadSecurityException.class, sheet.failure());
        }
    }
}
