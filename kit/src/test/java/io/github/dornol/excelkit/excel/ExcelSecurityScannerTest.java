package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadSecurityException;
import io.github.dornol.excelkit.core.ReadSecurityPolicy;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.ValueSource;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.nio.charset.StandardCharsets;
import java.util.function.UnaryOperator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipInputStream;
import java.util.zip.ZipOutputStream;

import static org.junit.jupiter.api.Assertions.*;

/** The formula policy must hold for any valid package layout, not only the default one POI writes. */
class ExcelSecurityScannerTest {
    private static final String SHEET = "xl/worksheets/sheet1.xml";

    private static byte[] workbook(boolean formula) throws Exception {
        try (var workbook = new XSSFWorkbook(); var out = new ByteArrayOutputStream()) {
            var row = workbook.createSheet("S").createRow(0);
            row.createCell(0).setCellValue(1);
            if (formula) row.createCell(1).setCellFormula("A1+1");
            workbook.write(out);
            return out.toByteArray();
        }
    }

    private static byte[] rewrite(byte[] source, UnaryOperator<String> names, UnaryOperator<String> text)
            throws Exception {
        var out = new ByteArrayOutputStream();
        try (var in = new ZipInputStream(new ByteArrayInputStream(source)); var zip = new ZipOutputStream(out)) {
            for (ZipEntry entry; (entry = in.getNextEntry()) != null; ) {
                String name = entry.getName();
                byte[] data = in.readAllBytes();
                if (name.endsWith(".xml") || name.endsWith(".rels")) {
                    data = text.apply(new String(data, StandardCharsets.UTF_8)).replace("\u0000", "")
                            .getBytes(StandardCharsets.UTF_8);
                }
                zip.putNextEntry(new ZipEntry(names.apply(name)));
                zip.write(data);
                zip.closeEntry();
            }
        }
        return out.toByteArray();
    }

    private static void read(byte[] file) {
        ExcelReader.forMap().securityPolicy(ReadSecurityPolicy.STRICT)
                .read(new ByteArrayInputStream(file), result -> { });
    }

    private static void assertFormulaRejected(byte[] file) {
        var error = assertThrows(ReadSecurityException.class, () -> read(file));
        assertEquals(ReadSecurityException.Reason.FORMULA, error.reason());
    }

    @Test
    void defaultLayoutFormulaIsRejectedAndPlainWorkbookIsAccepted() throws Exception {
        assertFormulaRejected(workbook(true));
        assertDoesNotThrow(() -> read(workbook(false)));
    }

    @Test
    void formulasWithANamespacePrefixAreRejected() throws Exception {
        byte[] file = rewrite(workbook(true), n -> n, text -> text.contains("<worksheet")
                ? text.replaceFirst("xmlns=\"", "xmlns:x=\"").replaceAll("<(/?)([A-Za-z])", "<$1x:$2") : text);
        assertFormulaRejected(file);
    }

    @ParameterizedTest
    @ValueSource(strings = {"XL/Worksheets/sheet1.xml", "xl/data/sheet1.xml", "custom/part.bin"})
    void formulasInAWorksheetPartAtAnyPathAreRejected(String path) throws Exception {
        byte[] file = rewrite(workbook(true), n -> n.equals(SHEET) ? path : n,
                text -> text.replace("/xl/worksheets/sheet1.xml", "/" + path)
                        .replace("worksheets/sheet1.xml", path.replaceFirst("(?i)^xl/", ""))
                        .replace("Target=\"/xl/", "Target=\"/xl/"));
        assertFormulaRejected(file);
    }

    @Test
    void extensionFormulaElementsAndChartReferencesAreNotCellFormulas() throws Exception {
        byte[] file = rewrite(workbook(false), n -> n, text -> text.contains("<worksheet")
                ? text.replace("</worksheet>", "<extLst><ext uri=\"{78C0D931-6437-407d-A8EE-F0AAD7539E65}\" "
                + "xmlns:x14=\"http://schemas.microsoft.com/office/spreadsheetml/2009/9/main\" "
                + "xmlns:xm=\"http://schemas.microsoft.com/office/excel/2006/main\"><x14:conditionalFormattings/>"
                + "<xm:f>1</xm:f></ext></extLst></worksheet>") : text);
        assertDoesNotThrow(() -> read(file));
    }
}
