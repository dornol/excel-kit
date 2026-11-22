package io.github.dornol.excelkit.core.read;

import io.github.dornol.excelkit.excel.read.ExcelReader;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.*;

class ReadSecurityPolicyEnforcementTest {

    @Test void strictSecurityPolicyRejectsFormulas() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var workbook = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
            var sheet = workbook.createSheet("Data");
            sheet.createRow(0).createCell(0).setCellValue("Value");
            sheet.createRow(1).createCell(0).setCellFormula("1+1");
            workbook.write(output);
        }
        ReadSecurityException error = assertThrows(ReadSecurityException.class, () ->
                ExcelReader.forMap().securityPolicy(ReadSecurityPolicy.STRICT)
                        .read(new ByteArrayInputStream(output.toByteArray()), result -> fail()));
        assertEquals(ReadSecurityException.Reason.FORMULA, error.reason());
    }

    @Test void strictSecurityPolicyRejectsExternalWorkbookRelationship() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var zip = new java.util.zip.ZipOutputStream(output)) {
            zip.putNextEntry(new java.util.zip.ZipEntry("xl/worksheets/_rels/sheet1.xml.rels"));
            zip.write("<Relationships><Relationship Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/externalLink\"/></Relationships>"
                    .getBytes(StandardCharsets.UTF_8));
            zip.closeEntry();
        }
        ReadSecurityException error = assertThrows(ReadSecurityException.class, () ->
                ExcelReader.forMap().securityPolicy(ReadSecurityPolicy.STRICT)
                        .read(new ByteArrayInputStream(output.toByteArray()), result -> fail()));
        assertEquals(ReadSecurityException.Reason.EXTERNAL_LINK, error.reason());
    }

    @Test void securityPolicyRejectsTooManyZipEntries() throws Exception {
        ByteArrayOutputStream output = new ByteArrayOutputStream();
        try (var zip = new java.util.zip.ZipOutputStream(output)) {
            for (int i = 0; i < 3; i++) {
                zip.putNextEntry(new java.util.zip.ZipEntry("entry-" + i));
                zip.write('x');
                zip.closeEntry();
            }
        }
        ReadSecurityPolicy policy = new ReadSecurityPolicy(true, true, 1024, 4096, 100.0, 2);
        ReadSecurityException error = assertThrows(ReadSecurityException.class, () ->
                ExcelReader.forMap().securityPolicy(policy)
                        .read(new ByteArrayInputStream(output.toByteArray()), result -> fail()));
        assertEquals(ReadSecurityException.Reason.ZIP_ENTRIES, error.reason());
    }
}
