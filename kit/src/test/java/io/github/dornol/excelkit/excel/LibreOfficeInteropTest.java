package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelColor;
import org.junit.jupiter.api.Assumptions;
import org.junit.jupiter.api.BeforeAll;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.List;
import java.util.concurrent.TimeUnit;
import java.util.stream.IntStream;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Opens what the writers produce in a real spreadsheet application. Excel is not available to the build; LibreOffice
 * is, in CI. A file that POI and the XML checks accept can still be one an application refuses (this is how the
 * ZIP64 data descriptors that POI 6 writes by default, which LibreOffice rejects, were found). The test is skipped
 * when no {@code soffice} is on the PATH, unless {@code REQUIRE_LIBREOFFICE=true} is set (the CI job sets it).
 */
class LibreOfficeInteropTest {

    private record Item(String name, int qty) {
    }

    private static String soffice;

    @BeforeAll
    static void findLibreOffice() {
        for (String dir : System.getenv().getOrDefault("PATH", "").split(java.io.File.pathSeparator)) {
            for (String name : new String[]{"soffice", "libreoffice"}) {
                if (soffice == null && Files.isExecutable(Path.of(dir, name))) soffice = Path.of(dir, name).toString();
            }
        }
        if (soffice == null && "true".equals(System.getenv("REQUIRE_LIBREOFFICE"))) {
            // CI installs LibreOffice for this test; not finding it must not turn into a skipped, green job
            org.junit.jupiter.api.Assertions.fail("REQUIRE_LIBREOFFICE is set but no soffice is on the PATH");
        }
        Assumptions.assumeTrue(soffice != null, "LibreOffice (soffice) is not installed; interop test skipped");
    }

    private static List<Item> items(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Item("item " + i, i)).toList();
    }

    /** Converts the workbook to one CSV per sheet and returns their lines, failing when LibreOffice cannot open it. */
    private static List<List<String>> openInLibreOffice(Path dir, String name, Path xlsx) throws Exception {
        Path profile = dir.resolve("profile-" + name);
        Process process = new ProcessBuilder(soffice, "-env:UserInstallation=" + profile.toUri(), "--headless",
                "--convert-to", "csv:Text - txt - csv (StarCalc):44,34,76,1,,0,false,true,false,false,false,-1",
                "--outdir", dir.resolve("out-" + name).toString(), xlsx.toString())
                .redirectErrorStream(true).start();
        String log = new String(process.getInputStream().readAllBytes(), StandardCharsets.UTF_8);
        assertTrue(process.waitFor(180, TimeUnit.SECONDS), "LibreOffice did not finish: " + log);
        assertEquals(0, process.exitValue(), log);
        assertTrue(!log.contains("could not be loaded"), name + ": LibreOffice cannot open the file: " + log);

        try (Stream<Path> files = Files.list(dir.resolve("out-" + name))) {
            // In the order of the sheets: "Sheet", "Sheet (2)", "Sheet (3)" do not sort that way by name ('.' follows ' ')
            return files.sorted(java.util.Comparator.comparingInt(LibreOfficeInteropTest::sheetNumber)
                    .thenComparing(f -> f.getFileName().toString())).map(f -> {
                try {
                    return Files.readAllLines(f, StandardCharsets.UTF_8);
                } catch (Exception e) {
                    throw new AssertionError(e);
                }
            }).toList();
        }
    }

    /** The number of a rolled-over sheet from the name of its file ("rich-Sheet (3).csv" is 3); the first sheet is 1. */
    private static int sheetNumber(Path file) {
        var m = java.util.regex.Pattern.compile("\\((\\d+)\\)\\.csv$").matcher(file.getFileName().toString());
        return m.find() ? Integer.parseInt(m.group(1)) : 1;
    }

    private static Path write(Path dir, String name, ExcelWriter<Item> writer, int rows) throws Exception {
        Path file = dir.resolve(name + ".xlsx");
        try (OutputStream out = Files.newOutputStream(file)) {
            writer.write(items(rows)).writeTo(out);
        }
        return file;
    }

    @Test
    void aPlainWorkbookOpensWithItsValues(@TempDir Path dir) throws Exception {
        Path file = write(dir, "plain", ExcelWriter.<Item>create().column("Name", Item::name).column("Qty", Item::qty), 5);

        List<List<String>> sheets = openInLibreOffice(dir, "plain", file);

        assertEquals(1, sheets.size());
        assertEquals(List.of("Name,Qty", "item 0,0", "item 1,1", "item 2,2", "item 3,3", "item 4,4"), sheets.get(0));
    }

    @Test
    void aWorkbookWithEveryStructureOpensAndKeepsItsRows(@TempDir Path dir) throws Exception {
        ExcelWriter<Item> writer = ExcelWriter.<Item>create().maxRows(4).table("InteropTable").autoFilter(true).freezeRows(1)
                .tabColor(ExcelColor.GREEN).protectSheet("lock")
                .column("Name", Item::name, c -> c.group("G"))
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER).group("G"))
                .conditionalFormatting(cf -> cf.columns(1).greaterThan("5", ExcelColor.LIGHT_RED))
                .summary(s -> s.sum("Qty").average("Qty"))
                .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).categoryColumn(0).valueColumn(1, "Qty"));
        Path file = write(dir, "rich", writer, 10);

        List<List<String>> sheets = openInLibreOffice(dir, "rich", file);

        assertEquals(3, sheets.size(), "10 rows with maxRows(4) fill three sheets");
        // group header row, column header row, then the data rows of the sheet
        assertEquals("item 0,0", sheets.get(0).get(2));
        assertEquals("item 3,3", sheets.get(0).get(5));
        assertEquals("item 4,4", sheets.get(1).get(2));
        assertEquals("item 8,8", sheets.get(2).get(2));
    }

    @Test
    void aMultiSheetWorkbookOpens(@TempDir Path dir) throws Exception {
        Path file = dir.resolve("multi.xlsx");
        try (ExcelWorkbook book = ExcelWorkbook.create(); OutputStream out = Files.newOutputStream(file)) {
            book.<Item>sheet("A").column("Name", Item::name).write(items(2).stream());
            book.<Item>sheet("B").column("Qty", Item::qty).write(items(3).stream());
            book.finish().writeTo(out);
        }

        List<List<String>> sheets = openInLibreOffice(dir, "multi", file);

        assertEquals(List.of(List.of("Name", "item 0", "item 1"), List.of("Qty", "0", "1", "2")), sheets);
    }
}
