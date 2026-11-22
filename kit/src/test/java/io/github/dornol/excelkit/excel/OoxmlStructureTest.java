package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.params.ParameterizedTest;
import org.junit.jupiter.params.provider.MethodSource;
import org.w3c.dom.Document;
import org.w3c.dom.Element;
import org.w3c.dom.Node;
import org.w3c.dom.NodeList;

import javax.xml.XMLConstants;
import javax.xml.parsers.DocumentBuilderFactory;
import javax.xml.xpath.XPathConstants;
import javax.xml.xpath.XPathFactory;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.nio.charset.StandardCharsets;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.HashSet;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.function.Supplier;
import java.util.stream.IntStream;
import java.util.zip.ZipEntry;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertFalse;
import static org.junit.jupiter.api.Assertions.assertNotNull;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Checks the {@code .xlsx} package with the JDK's XML parser only, never with POI. POI is both what writes the
 * file and what the other tests read it with, so a defect POI forgives on both sides would go unnoticed; Excel
 * and LibreOffice are less forgiving. These are the structural rules a spreadsheet application relies on
 * (every part referenced exists, style indexes resolve, rows and cells are ordered, merges do not overlap)
 * plus the exact XML a feature is expected to produce.
 */
class OoxmlStructureTest {

    private record Item(String name, int qty, double price, LocalDate day, String url) {
    }

    private static List<Item> items(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Item("item " + i, i, i * 1.5,
                LocalDate.of(2024, 1, 1).plusDays(i), "https://example.com/" + i)).toList();
    }

    // ---- a minimal view of the package ------------------------------------------------------------------

    static final class Pkg {
        final Map<String, byte[]> parts = new java.util.LinkedHashMap<>();
        final Map<String, Document> xml = new HashMap<>();

        Pkg(byte[] xlsx) throws Exception {
            // ZipInputStream cannot read a streamed (data-descriptor) archive, so go through a file.
            java.nio.file.Path tmp = java.nio.file.Files.createTempFile("ooxml-structure", ".xlsx");
            try {
                java.nio.file.Files.write(tmp, xlsx);
                try (java.util.zip.ZipFile zip = new java.util.zip.ZipFile(tmp.toFile())) {
                    for (var entries = zip.entries(); entries.hasMoreElements(); ) {
                        ZipEntry e = entries.nextElement();
                        try (var in = zip.getInputStream(e)) {
                            parts.put(e.getName(), in.readAllBytes());
                        }
                    }
                }
            } finally {
                java.nio.file.Files.deleteIfExists(tmp);
            }
            DocumentBuilderFactory f = DocumentBuilderFactory.newInstance();
            f.setNamespaceAware(true);
            f.setFeature(XMLConstants.FEATURE_SECURE_PROCESSING, true);
            f.setFeature("http://apache.org/xml/features/disallow-doctype-decl", true);
            for (var e : parts.entrySet()) {
                if (e.getKey().endsWith(".xml") || e.getKey().endsWith(".rels") || e.getKey().endsWith(".vml")) {
                    // A part that is not well-formed fails here, naming the part.
                    try {
                        xml.put(e.getKey(), f.newDocumentBuilder().parse(new ByteArrayInputStream(e.getValue())));
                    } catch (Exception ex) {
                        throw new AssertionError("part " + e.getKey() + " is not well-formed XML: " + ex.getMessage(), ex);
                    }
                }
            }
        }

        Document doc(String part) {
            Document d = xml.get(part);
            assertNotNull(d, "missing part " + part + " in " + parts.keySet());
            return d;
        }

        List<Element> select(String part, String xpath) {
            try {
                NodeList nodes = (NodeList) XPathFactory.newInstance().newXPath()
                        .evaluate(xpath, doc(part), XPathConstants.NODESET);
                List<Element> out = new ArrayList<>();
                for (int i = 0; i < nodes.getLength(); i++) out.add((Element) nodes.item(i));
                return out;
            } catch (Exception e) {
                throw new AssertionError(e);
            }
        }

        List<String> sheetParts() {
            return parts.keySet().stream().filter(n -> n.matches("xl/worksheets/sheet\\d+\\.xml"))
                    .sorted(java.util.Comparator.comparingInt(n -> Integer.parseInt(n.replaceAll("\\D", "")))).toList();
        }
    }

    private static final String MAIN = "//*[local-name()='%s']";

    static String local(String name) {
        return String.format(MAIN, name);
    }

    private static byte[] write(ExcelWriter<Item> writer, int rows) throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        writer.write(items(rows)).writeTo(out);
        return out.toByteArray();
    }

    // ---- invariants that hold for every workbook the library writes ---------------------------------------

    static void assertValidPackage(Pkg pkg) {
        assertTrue(pkg.parts.containsKey("[Content_Types].xml"), "no [Content_Types].xml");
        assertTrue(pkg.parts.containsKey("_rels/.rels"), "no root relationships");

        // Content types: every Override names an existing part; every part has a type.
        Set<String> defaults = new HashSet<>();
        for (Element d : pkg.select("[Content_Types].xml", local("Default"))) defaults.add(d.getAttribute("Extension").toLowerCase());
        Set<String> overrides = new HashSet<>();
        for (Element o : pkg.select("[Content_Types].xml", local("Override"))) {
            String part = o.getAttribute("PartName").substring(1);
            overrides.add(part);
            assertTrue(pkg.parts.containsKey(part), "content type for a part that does not exist: " + part);
        }
        for (String part : pkg.parts.keySet()) {
            if (part.equals("[Content_Types].xml")) continue;
            String ext = part.substring(part.lastIndexOf('.') + 1).toLowerCase();
            assertTrue(overrides.contains(part) || defaults.contains(ext), "part without a content type: " + part);
        }

        // Relationships: every internal target exists.
        for (String part : pkg.parts.keySet()) {
            if (!part.endsWith(".rels")) continue;
            String base = part.equals("_rels/.rels") ? "" : part.replace("_rels/", "").replaceAll("[^/]*\\.rels$", "");
            for (Element rel : pkg.select(part, local("Relationship"))) {
                if ("External".equals(rel.getAttribute("TargetMode"))) continue;
                String target = rel.getAttribute("Target");
                String resolved = target.startsWith("/") ? target.substring(1)
                        : java.net.URI.create("/" + base).resolve(target).getPath().substring(1);
                assertTrue(pkg.parts.containsKey(resolved), part + " points at " + target + " (" + resolved + ")");
            }
        }

        // Styles: every index a cell uses resolves.
        List<Element> xfs = pkg.select("xl/styles.xml", "//*[local-name()='cellXfs']/*[local-name()='xf']");
        int fonts = pkg.select("xl/styles.xml", "//*[local-name()='fonts']/*[local-name()='font']").size();
        int fills = pkg.select("xl/styles.xml", "//*[local-name()='fills']/*[local-name()='fill']").size();
        int borders = pkg.select("xl/styles.xml", "//*[local-name()='borders']/*[local-name()='border']").size();
        Set<Integer> customFormats = new HashSet<>();
        for (Element n : pkg.select("xl/styles.xml", "//*[local-name()='numFmts']/*[local-name()='numFmt']")) {
            customFormats.add(Integer.parseInt(n.getAttribute("numFmtId")));
        }
        for (Element xf : xfs) {
            assertTrue(Integer.parseInt(xf.getAttribute("fontId")) < fonts, "xf font out of range");
            assertTrue(Integer.parseInt(xf.getAttribute("fillId")) < fills, "xf fill out of range");
            assertTrue(Integer.parseInt(xf.getAttribute("borderId")) < borders, "xf border out of range");
            int fmt = Integer.parseInt(xf.getAttribute("numFmtId"));
            assertTrue(fmt < 164 || customFormats.contains(fmt), "xf uses undefined number format " + fmt);
        }
        List<Element> sharedStrings = pkg.parts.containsKey("xl/sharedStrings.xml")
                ? pkg.select("xl/sharedStrings.xml", "//*[local-name()='si']") : List.of();

        // Sheets.
        Set<String> names = new HashSet<>();
        Set<String> ids = new HashSet<>();
        for (Element s : pkg.select("xl/workbook.xml", local("sheet"))) {
            String name = s.getAttribute("name");
            assertTrue(name.length() <= 31 && !name.isEmpty(), "bad sheet name <" + name + ">");
            assertTrue(names.add(name.toLowerCase()), "duplicate sheet name " + name);
            assertTrue(ids.add(s.getAttribute("sheetId")), "duplicate sheetId " + s.getAttribute("sheetId"));
        }
        assertEquals(names.size(), pkg.sheetParts().size(), "sheet entries vs worksheet parts");

        for (String sheet : pkg.sheetParts()) {
            int lastRow = 0;
            for (Element row : pkg.select(sheet, "//*[local-name()='sheetData']/*[local-name()='row']")) {
                int r = Integer.parseInt(row.getAttribute("r"));
                assertTrue(r > lastRow, sheet + ": row " + r + " after " + lastRow);
                lastRow = r;
                int lastCol = 0;
                for (Node n = row.getFirstChild(); n != null; n = n.getNextSibling()) {
                    if (!(n instanceof Element c) || !c.getLocalName().equals("c")) continue;
                    String ref = c.getAttribute("r");
                    assertTrue(ref.endsWith(String.valueOf(r)), sheet + ": cell " + ref + " in row " + r);
                    int col = columnNumber(ref);
                    assertTrue(col > lastCol, sheet + ": cell " + ref + " not after column " + lastCol);
                    lastCol = col;
                    if (c.hasAttribute("s")) {
                        assertTrue(Integer.parseInt(c.getAttribute("s")) < xfs.size(), sheet + ": style index out of range at " + ref);
                    }
                    if ("s".equals(c.getAttribute("t"))) {
                        String idx = c.getElementsByTagNameNS("*", "v").item(0).getTextContent();
                        assertTrue(Integer.parseInt(idx) < sharedStrings.size(), sheet + ": shared string out of range at " + ref);
                    }
                }
            }
            assertMergesDoNotOverlap(pkg, sheet);
        }
    }

    private static boolean isTrue(String xsdBoolean) {
        return xsdBoolean.equals("1") || xsdBoolean.equals("true");
    }

    private static int columnNumber(String ref) {
        int col = 0;
        for (char ch : ref.toCharArray()) {
            if (ch < 'A' || ch > 'Z') break;
            col = col * 26 + (ch - 'A' + 1);
        }
        return col;
    }

    private static int[] range(String ref) {
        String[] ends = ref.split(":");
        String[] a = ends[0].split("(?<=[A-Z])(?=\\d)");
        String[] b = (ends.length > 1 ? ends[1] : ends[0]).split("(?<=[A-Z])(?=\\d)");
        return new int[]{Integer.parseInt(a[1]), columnNumber(a[0]), Integer.parseInt(b[1]), columnNumber(b[0])};
    }

    private static void assertMergesDoNotOverlap(Pkg pkg, String sheet) {
        List<Element> merges = pkg.select(sheet, local("mergeCell"));
        List<Element> container = pkg.select(sheet, local("mergeCells"));
        if (!container.isEmpty()) {
            assertEquals(merges.size(), Integer.parseInt(container.get(0).getAttribute("count")), sheet + ": mergeCells count");
        }
        List<int[]> ranges = merges.stream().map(m -> range(m.getAttribute("ref"))).toList();
        for (int i = 0; i < ranges.size(); i++) {
            for (int j = i + 1; j < ranges.size(); j++) {
                int[] x = ranges.get(i), y = ranges.get(j);
                boolean overlap = x[0] <= y[2] && y[0] <= x[2] && x[1] <= y[3] && y[1] <= x[3];
                assertFalse(overlap, sheet + ": merged ranges overlap " + merges.get(i).getAttribute("ref")
                        + " and " + merges.get(j).getAttribute("ref"));
            }
        }
    }

    // ---- workbooks of many shapes all satisfy the invariants -----------------------------------------------

    private static List<Object[]> shapes() {
        List<Object[]> shapes = new ArrayList<>();
        shapes.add(new Object[]{"plain", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().column("Name", Item::name).column("Qty", Item::qty), 5))});
        shapes.add(new Object[]{"empty", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().column("Name", Item::name), 0))});
        shapes.add(new Object[]{"rollover", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().maxRows(7).column("Name", Item::name).rowNumberColumn("No."), 40))});
        shapes.add(new Object[]{"rollover with every sheet feature", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().maxRows(10).autoFilter(true).freezeRows(1).tabColor(ExcelColor.GREEN)
                        .column("Name", Item::name, c -> c.group("G"))
                        .column("Qty", Item::qty, c -> c.group("G"))
                        .conditionalFormatting(cf -> cf.columns(1).greaterThan("5", ExcelColor.LIGHT_RED)), 35))});
        shapes.add(new Object[]{"multi-level groups", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create()
                        .column("A", Item::qty, c -> c.group("Top", "Mid"))
                        .column("B", Item::qty, c -> c.group("Top", "Mid"))
                        .column("C", Item::qty, c -> c.group("Top"))
                        .column("D", Item::qty), 4))});
        shapes.add(new Object[]{"table", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().table("ItemsTable").column("Name", Item::name).column("Qty", Item::qty), 12))});
        shapes.add(new Object[]{"table with rollover", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().maxRows(5).table("Rolled").column("Name", Item::name).column("Qty", Item::qty), 12))});
        shapes.add(new Object[]{"styles and formats", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create()
                        .column("Day", Item::day)
                        .column("Price", Item::price, c -> c.type(ExcelDataType.DOUBLE).bold(true))
                        .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER).cellColor((v, r) ->
                                ((Number) v).intValue() % 2 == 0 ? ExcelColor.LIGHT_YELLOW : null)), 20))});
        shapes.add(new Object[]{"summary", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                        .summary(s -> s.label("Total").sum("Qty")), 9))});
        shapes.add(new Object[]{"hyperlinks, validation, protection", (Supplier<byte[]>) () -> sneaky(() -> write(
                ExcelWriter.<Item>create().protectSheet("pw")
                        .column("Url", Item::url, c -> c.type(ExcelDataType.HYPERLINK))
                        .column("Name", Item::name, c -> c.dropdown("a", "b").locked(false)), 6))});
        return shapes;
    }

    private interface IoSupplier {
        byte[] get() throws Exception;
    }

    private static byte[] sneaky(IoSupplier s) {
        try {
            return s.get();
        } catch (Exception e) {
            throw new AssertionError(e);
        }
    }

    @ParameterizedTest(name = "{0}")
    @MethodSource("shapes")
    void everyWorkbookShapeIsAStructurallyValidPackage(String name, Supplier<byte[]> workbook) throws Exception {
        assertValidPackage(new Pkg(workbook.get()));
    }

    // ---- the exact XML of individual features ------------------------------------------------------------

    @Test
    void freezeAutoFilterAndTabColorAreWrittenWhereExcelLooksForThem() throws Exception {
        Pkg pkg = new Pkg(write(ExcelWriter.<Item>create().freezePane(2, 1).autoFilter(true).tabColor(255, 0, 0)
                .column("Name", Item::name).column("Qty", Item::qty).column("Price", Item::price), 5));

        Element pane = pkg.select("xl/worksheets/sheet1.xml", local("pane")).get(0);
        assertEquals(2.0, Double.parseDouble(pane.getAttribute("xSplit"))); // xsd:double
        assertEquals(1.0, Double.parseDouble(pane.getAttribute("ySplit")));
        assertEquals("frozen", pane.getAttribute("state"));
        assertEquals("C2", pane.getAttribute("topLeftCell"));

        // The filter covers the header row only (the data is streamed, so the row count is not known when it is
        // set); whether Excel/LibreOffice extend it over the data on filtering is not verified here.
        assertEquals("A1:C1", pkg.select("xl/worksheets/sheet1.xml", local("autoFilter")).get(0).getAttribute("ref"));
        String tab = pkg.select("xl/worksheets/sheet1.xml", local("tabColor")).get(0).getAttribute("rgb");
        assertTrue(tab.equals("FF0000") || tab.equals("FFFF0000"), "tab colour " + tab); // alpha is optional
        // autoFilter needs a hidden built-in name or Excel drops the filter on open
        List<Element> names = pkg.select("xl/workbook.xml", local("definedName"));
        assertTrue(names.stream().anyMatch(n -> n.getAttribute("name").equals("_xlnm._FilterDatabase")),
                "no _FilterDatabase defined name: " + names.size());
        assertValidPackage(pkg);
    }

    @Test
    void groupHeadersMergeExactlyTheDeclaredRanges() throws Exception {
        Pkg pkg = new Pkg(write(ExcelWriter.<Item>create()
                .column("A", Item::qty, c -> c.group("Top", "Mid"))
                .column("B", Item::qty, c -> c.group("Top", "Mid"))
                .column("C", Item::qty, c -> c.group("Top"))
                .column("D", Item::qty), 2));

        Set<String> merges = new HashSet<>();
        for (Element m : pkg.select("xl/worksheets/sheet1.xml", local("mergeCell"))) merges.add(m.getAttribute("ref"));
        assertTrue(merges.contains("A1:C1"), "Top spans A-C on level 1: " + merges);
        assertTrue(merges.contains("A2:B2"), "Mid spans A-B on level 2: " + merges);
        assertTrue(merges.contains("C2:C3"), "C merges down into the column header: " + merges);
        assertTrue(merges.contains("D1:D3"), "D has no group and spans the whole header: " + merges);
        assertEquals(4, merges.size(), merges.toString());
    }

    @Test
    void dataValidationAndConditionalFormattingCarryTheirRangesAndRules() throws Exception {
        Pkg pkg = new Pkg(write(ExcelWriter.<Item>create()
                .column("Name", Item::name, c -> c.dropdown("Active", "Inactive"))
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .conditionalFormatting(cf -> cf.columns(1).greaterThan("10", ExcelColor.LIGHT_RED)), 3));
        String sheet = "xl/worksheets/sheet1.xml";

        Element validation = pkg.select(sheet, local("dataValidation")).get(0);
        assertEquals("list", validation.getAttribute("type"));
        assertTrue(validation.getAttribute("sqref").startsWith("A2"), validation.getAttribute("sqref"));
        String formula = validation.getElementsByTagNameNS("*", "formula1").item(0).getTextContent();
        assertEquals("\"Active,Inactive\"", formula);

        Element cf = pkg.select(sheet, local("conditionalFormatting")).get(0);
        assertTrue(cf.getAttribute("sqref").startsWith("B2"), cf.getAttribute("sqref"));
        Element rule = (Element) cf.getElementsByTagNameNS("*", "cfRule").item(0);
        assertEquals("cellIs", rule.getAttribute("type"));
        assertEquals("greaterThan", rule.getAttribute("operator"));
        assertEquals("10", rule.getElementsByTagNameNS("*", "formula").item(0).getTextContent());
        // the rule's colour lives in a differential style, which has to exist
        int dxfId = Integer.parseInt(rule.getAttribute("dxfId"));
        assertTrue(dxfId < pkg.select("xl/styles.xml", local("dxf")).size(), "dxfId " + dxfId + " has no style");
    }

    @Test
    void sheetProtectionAndHiddenColumnsAreMarked() throws Exception {
        Pkg pkg = new Pkg(write(ExcelWriter.<Item>create().protectSheet("pw")
                .column("Name", Item::name, c -> c.locked(false))
                .column("Qty", Item::qty, c -> c.hidden()), 2));
        String sheet = "xl/worksheets/sheet1.xml";

        Element protection = pkg.select(sheet, local("sheetProtection")).get(0);
        assertTrue(isTrue(protection.getAttribute("sheet")), "sheet=" + protection.getAttribute("sheet")); // xsd:boolean
        assertFalse(protection.getAttribute("hashValue").isEmpty() && protection.getAttribute("password").isEmpty(),
                "the protection must carry a password hash, not be open");
        List<Element> cols = pkg.select(sheet, local("col"));
        assertTrue(cols.stream().anyMatch(c -> isTrue(c.getAttribute("hidden")) && c.getAttribute("min").equals("2")),
                "column B hidden");
        assertTrue(cols.stream().noneMatch(c -> isTrue(c.getAttribute("hidden")) && c.getAttribute("min").equals("1")),
                "column A visible");
        // An unlocked column needs an xf with locked=0.
        assertTrue(pkg.select("xl/styles.xml", "//*[local-name()='cellXfs']//*[local-name()='protection']").stream()
                .anyMatch(p -> !isTrue(p.getAttribute("locked"))), "no unlocked style");
    }

    @Test
    void tablePartsAreWiredToTheirSheetsWithUniqueNamesAndIds() throws Exception {
        Pkg pkg = new Pkg(write(ExcelWriter.<Item>create().maxRows(5).table("Rolled")
                .column("Name", Item::name).column("Qty", Item::qty), 12));

        List<String> tables = pkg.parts.keySet().stream().filter(n -> n.matches("xl/tables/table\\d+\\.xml")).sorted().toList();
        assertEquals(3, tables.size(), "one table per rolled sheet: " + tables);
        Set<String> tableNames = new HashSet<>();
        Set<String> tableIds = new HashSet<>();
        for (String table : tables) {
            Element t = pkg.doc(table).getDocumentElement();
            assertTrue(tableNames.add(t.getAttribute("name")), "duplicate table name " + t.getAttribute("name"));
            assertTrue(tableIds.add(t.getAttribute("id")), "duplicate table id " + t.getAttribute("id"));
            assertEquals(2, pkg.select(table, local("tableColumn")).size());
        }
        // each sheet declares its table part
        for (String sheet : pkg.sheetParts()) {
            assertEquals(1, pkg.select(sheet, local("tablePart")).size(), sheet + " has no tablePart");
        }
        assertValidPackage(pkg);
    }

    @Test
    void hyperlinksPointAtTheirTargets() throws Exception {
        Pkg pkg = new Pkg(write(ExcelWriter.<Item>create()
                .column("Url", Item::url, c -> c.type(ExcelDataType.HYPERLINK)), 3));

        List<Element> links = pkg.select("xl/worksheets/sheet1.xml", local("hyperlink"));
        assertEquals(3, links.size());
        Set<String> targets = new HashSet<>();
        for (Element rel : pkg.select("xl/worksheets/_rels/sheet1.xml.rels", local("Relationship"))) {
            if (rel.getAttribute("Type").endsWith("/hyperlink")) {
                assertEquals("External", rel.getAttribute("TargetMode"));
                targets.add(rel.getAttribute("Target"));
            }
        }
        assertEquals(Set.of("https://example.com/0", "https://example.com/1", "https://example.com/2"), targets);
    }

    @Test
    void textWithMarkupAndControlCharactersStaysWellFormed() throws Exception {
        // These are the characters that break XML if the writer does not escape or strip them.
        String nasty = "a<b>&\"'c \u0001\u0008 😀 ]]> end";
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create().column("T", s -> s).write(List.of(nasty)).writeTo(out);

        Pkg pkg = new Pkg(out.toByteArray()); // throws if any part is not well-formed
        String all = new String(pkg.parts.get("xl/sharedStrings.xml") != null ? pkg.parts.get("xl/sharedStrings.xml")
                : pkg.parts.get("xl/worksheets/sheet1.xml"), StandardCharsets.UTF_8);
        assertFalse(all.contains("\u0001"), "a raw control character reached the XML");
        assertValidPackage(pkg);
    }
}
