package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.ReadResult;
import io.github.dornol.excelkit.core.ReadSecurityException;
import io.github.dornol.excelkit.core.ReadSecurityPolicy;
import io.github.dornol.excelkit.core.RowData;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.openxml4j.opc.PackageAccess;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.HexFormat;
import java.util.List;
import java.util.Locale;
import java.util.Objects;
import java.util.zip.ZipFile;
import static org.junit.jupiter.api.Assertions.*;

/** Frozen upstream files, never produced or rewritten by this library during tests. */
class ExternalWorkbookFixtureTest {
    @BeforeEach
    void fixedFormattingLocale() {
        LocaleUtil.setUserLocale(Locale.US);
    }

    @Test
    void fixtureChecksumsAndProducerMetadataMatchThePinnedSources() throws Exception {
        verifyFixtureChecksums();
        assertProducer("WithVariousData.xlsx", "Microsoft Excel");
        assertProducer("DateFormatTests.xlsx", "Microsoft Macintosh Excel");
        assertProducer("shared_formulas.xlsx", "Microsoft Excel");
        assertProducer("59775.xlsx", "LibreOffice/5.0.3.2");
        assertProducer("70005-countifs.xlsx", "LibreOffice/25.2.7.2");
        assertProducer("60384.xlsx", "LibreOffice/5.1.4.2");
        assertProducer("tile-range-test.xlsx", "LibreOffice/5.3.4.2");
    }

    @Test
    void excelSharedStringsSparseRowsAndSheetMetadataAreReadWithoutRewriting() throws Exception {
        Path file = fixture("WithVariousData.xlsx");
        assertEquals(List.of("Sheet1", "Sheet2", "Sheet3"), ExcelReader.getSheetNames(file).stream().map(ExcelSheetInfo::name).toList());
        assertEquals(List.of("A1", "B1", "C1"), ExcelReader.getSheetHeaders(file, 0, 0));
        var rows = read(file, 0, null);
        assertEquals(8, rows.size());
        assertEquals("22.3", rows.get(0).data().get(0).asString());
        assertEquals("Foo", rows.get(0).data().get("B1").asString());
        assertEquals("I have two comments", atRow(rows, 7).get(2).asString());
        assertEquals("We have a header", atRow(rows, 9).get(2).asString());
        assertEquals("We have a footer", atRow(rows, 10).get(2).asString());
    }

    @Test
    void macExcel1904DatesAgreeWithCachedTextResultsAndHeaderExtraction() throws Exception {
        Path file = fixture("DateFormatTests.xlsx");
        var rows = read(file, 1, null);
        assertEquals(45, rows.size());
        assertEquals("11-10-52", rows.get(0).data().get(0).asString());
        assertTrue(rows.get(0).data().get(2).asString().contains("1952"));
        assertTrue(ExcelReader.getSheetHeaders(file, 1, 1).get(2).contains("1952"));
        try (var pkg = OPCPackage.open(file.toFile(), PackageAccess.READ)) {
            assertEquals("1952-10-11 14:35:27", ExcelReadSupport.dataFormatter(new XSSFReader(pkg))
                    .formatRawCellContents(17816.607951388887, 164, "yyyy-MM-dd HH:mm:ss"));
        }
        // A separate 1900 workbook must not inherit the previous file's date system.
        try (var pkg = OPCPackage.open(fixture("WithVariousData.xlsx").toFile(), PackageAccess.READ)) {
            assertEquals("1948-10-10 14:35:27", ExcelReadSupport.dataFormatter(new XSSFReader(pkg))
                    .formatRawCellContents(17816.607951388887, 164, "yyyy-MM-dd HH:mm:ss"));
        }
    }

    @Test
    void excelSharedFormulaCachedStringsArePreserved() throws Exception {
        var rows = read(fixture("shared_formulas.xlsx"), 0, null);
        assertEquals(40, rows.size());
        assertEquals("ProductionOrderConfirmation", atRow(rows, 2).get(0).asString());
        assertEquals("ProductionOrderConfirmation", atRow(rows, 2).get(1).asString());
        assertEquals("Code", atRow(rows, 41).get(0).asString());
    }

    @Test
    void libreOfficeHyperlinksAndMissingPhysicalRowsArePreservedAsText() throws Exception {
        var rows = read(fixture("59775.xlsx"), 0, null);
        assertEquals(List.of(2L, 3L, 4L, 6L, 7L), rows.stream().map(ReadResult::fileRowNum).toList());
        assertEquals("http://twitter.com/#!/apacheorg", rows.get(0).data().get("URLs").asString());
        assertEquals("https://en.wikipedia.org/wiki/Apache_POI#See_also", atRow(rows, 4).get(0).asString());
        assertEquals("#Sheet1", atRow(rows, 7).get(0).asString());
    }

    @Test
    void libreOfficeMergedTextAndFormattedNumbersSurviveTheStreamingReader() throws Exception {
        var rows = read(fixture("60384.xlsx"), 0, null);
        assertEquals(6, rows.size());
        assertEquals("{Parameter}", atRow(rows, 3).get(0).asString());
        assertEquals("0.10", atRow(rows, 3).get(1).asString());
        assertEquals("Titel", atRow(rows, 8).get(0).asString());
        assertEquals("An very very long template text reaching over multiple columns", atRow(rows, 9).get(0).asString());
    }

    @Test
    void libreOfficeCachedNumericFormulasWorkThroughSharedPackageReads() throws Exception {
        var first = new ArrayList<ReadResult<RowData>>();
        var second = new ArrayList<ReadResult<RowData>>();
        var report = ExcelWorkbookReader.create()
                .sheet("Sheet1", ExcelReader.<RowData>mapping(row -> row).skipBlankRows(), first::add)
                .sheet("Sheet2", ExcelReader.<RowData>mapping(row -> row).skipBlankRows(), second::add)
                .read(fixture("tile-range-test.xlsx"));
        assertTrue(report.successful());
        assertTrue(first.stream().allMatch(row -> row.success()));
        assertTrue(second.stream().allMatch(row -> row.success()));
        assertEquals(5, first.size());
        assertEquals(3, second.size());
        assertEquals("-0.2", atRow(first, 4).get(1).asString());
        assertEquals("3.1", atRow(first, 5).get(1).asString());
        assertEquals("3.4", atRow(first, 5).get(2).asString());
        assertEquals("21", atRow(second, 2).get(0).asString());
        var error = assertThrows(ReadSecurityException.class, () -> ExcelReader.forMap()
                .securityPolicy(ReadSecurityPolicy.STRICT).read(fixture("tile-range-test.xlsx"), row -> fail()));
        assertEquals(ReadSecurityException.Reason.FORMULA, error.reason());
    }

    @Test
    void modernLibreOfficeArrayFormulaCachesAreReadWithoutRecalculation() throws Exception {
        var rows = read(fixture("70005-countifs.xlsx"), 0, null);
        assertEquals(List.of(3L, 4L), rows.stream().map(ReadResult::fileRowNum).toList());
        assertEquals("3", rows.get(0).data().get(0).asString());
        assertEquals("3", rows.get(0).data().get(3).asString());
        assertEquals("4", rows.get(1).data().get(0).asString());
        assertEquals("4", rows.get(1).data().get(3).asString());
    }

    @AfterEach
    void fixtureBytesRemainUnchanged() throws Exception {
        try { verifyFixtureChecksums(); }
        finally { LocaleUtil.resetUserLocale(); }
    }

    private static void verifyFixtureChecksums() throws Exception {
        var checksums = Files.readAllLines(fixture("SHA256SUMS"), StandardCharsets.UTF_8);
        assertEquals(8, checksums.size());
        for (String line : checksums) {
            var fields = line.split("  ", 2);
            assertEquals(fields[0], HexFormat.of().formatHex(MessageDigest.getInstance("SHA-256")
                    .digest(Files.readAllBytes(fixture(fields[1])))), fields[1]);
        }
    }

    private static List<ReadResult<RowData>> read(Path path, int sheet, String password) {
        var rows = new ArrayList<ReadResult<RowData>>();
        var reader = ExcelReader.<RowData>mapping(row -> row).sheetIndex(sheet).skipBlankRows();
        if (password != null) reader.password(password);
        reader.read(path, row -> { assertTrue(row.success(), row.messages().toString()); rows.add(row); });
        return rows;
    }

    private static RowData atRow(List<ReadResult<RowData>> rows, long physicalRow) {
        return rows.stream().filter(row -> row.fileRowNum() == physicalRow).findFirst().orElseThrow().data();
    }

    private static void assertProducer(String name, String application) throws Exception {
        try (var zip = new ZipFile(fixture(name).toFile()); var input = zip.getInputStream(zip.getEntry("docProps/app.xml"))) {
            var document = XMLHelper.newDocumentBuilder().parse(input);
            assertTrue(document.getElementsByTagNameNS("*", "Application").item(0).getTextContent().startsWith(application));
        }
    }

    private static Path fixture(String name) throws Exception {
        return Path.of(Objects.requireNonNull(ExternalWorkbookFixtureTest.class.getResource("/external-workbooks/" + name)).toURI());
    }
}
