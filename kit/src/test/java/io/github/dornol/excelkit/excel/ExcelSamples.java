package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelImage;
import io.github.dornol.excelkit.excel.style.ExcelColor;

import javax.imageio.ImageIO;
import java.awt.Color;
import java.awt.Graphics2D;
import java.awt.image.BufferedImage;
import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.ArrayList;
import java.util.List;
import java.util.stream.IntStream;

/**
 * Writes the workbooks that only a real Excel can judge (does it open without a repair prompt, does a filter reach the
 * data, is a chart where it should be, ...) together with {@code manifest.json}, which says for each file what Excel
 * must show, and {@code CHECKLIST.md}, the same for a person. Run with {@code ./gradlew :kit:excelSamples}; the files
 * are written to {@code kit/build/excel-samples}. {@link ExcelSamplesTest} keeps the manifest true as far as POI can tell.
 */
public final class ExcelSamples {

    private record Item(String name, int qty) {
    }

    /** One sample: the file, what to look at, and checks a program can make through Excel's object model. */
    record Sample(String file, String title, String look, List<String> checks) {
    }

    private final Path dir;
    private final List<Sample> samples = new ArrayList<>();

    private ExcelSamples(Path dir) {
        this.dir = dir;
    }

    public static void main(String[] args) throws Exception {
        Path dir = Path.of(args.length > 0 ? args[0] : "build/excel-samples");
        generate(dir);
        System.out.println("Wrote " + Files.list(dir).count() + " files to " + dir.toAbsolutePath());
    }

    static List<Sample> generate(Path dir) throws Exception {
        Files.createDirectories(dir);
        ExcelSamples s = new ExcelSamples(dir);
        s.write();
        Files.writeString(dir.resolve("manifest.json"), s.manifest(), StandardCharsets.UTF_8);
        Files.writeString(dir.resolve("CHECKLIST.md"), s.checklist(), StandardCharsets.UTF_8);
        return s.samples;
    }

    // ---- helpers -----------------------------------------------------------------------------------------

    private static List<Item> items(int n) {
        return IntStream.range(0, n).mapToObj(i -> new Item("item " + i, i)).toList();
    }

    private static ExcelWriter<Item> report() {
        return ExcelWriter.<Item>create().sheetName("Report").column("Name", Item::name)
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER));
    }

    private void save(String file, ExcelHandler handler) throws IOException {
        try (OutputStream out = Files.newOutputStream(dir.resolve(file))) {
            handler.writeTo(out);
        }
    }

    private void add(String file, String title, String look, String... checks) {
        samples.add(new Sample(file, title, look, List.of(checks)));
    }

    /** A JSON string literal of the given text. */
    private static String q(String text) {
        StringBuilder sb = new StringBuilder("\"");
        for (char c : text.toCharArray()) {
            switch (c) {
                case '"' -> sb.append("\\\"");
                case '\\' -> sb.append("\\\\");
                case '\n' -> sb.append("\\n");
                default -> sb.append(c);
            }
        }
        return sb.append('"').toString();
    }

    private static byte[] png(Color color) throws IOException {
        BufferedImage image = new BufferedImage(48, 48, BufferedImage.TYPE_INT_RGB);
        Graphics2D g = image.createGraphics();
        g.setColor(color);
        g.fillRect(0, 0, 48, 48);
        g.setColor(Color.WHITE);
        g.fillOval(12, 12, 24, 24);
        g.dispose();
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ImageIO.write(image, "png", out);
        return out.toByteArray();
    }

    // ---- the samples -------------------------------------------------------------------------------------

    private void write() throws Exception {
        // 1. a filter over header and data, summary rows outside it
        save("01-autofilter-plain.xlsx", report().autoFilter(true).summary(s -> s.sum("Qty").average("Qty")).write(items(7)));
        add("01-autofilter-plain.xlsx", "autoFilter over the data",
                "Filter buttons on the header row (A1:B1). Filter Name = \"item 3\": ONE data row stays, the rows below the data (Sum, Average) are not hidden.",
                "opens: no repair prompt",
                "autoFilter sheet=Report range=A1:B8 (the range in the file; Excel is known to report A1:B10 here, see sample 17: report the range Excel shows right after opening, before any filter call)",
                "filterRows sheet=Report column=1 criteria=\"item 3\" expectVisibleDataRows=1 (data range A2:A8); Excel is known to hide Sum and Average as well, which is what sample 17 is about",
                "cell sheet=Report cell=A9 text=Sum; cell sheet=Report cell=B9 value=21; cell sheet=Report cell=A10 text=Average; cell sheet=Report cell=B10 value=3");

        // 2. grouped header: the filter starts at the row of the column names
        save("02-autofilter-group-header.xlsx", ExcelWriter.<Item>create().sheetName("Report").autoFilter(true)
                .column("Name", Item::name, c -> c.group("G")).column("Qty", Item::qty, c -> c.group("G")).write(items(3)));
        add("02-autofilter-group-header.xlsx", "autoFilter under a group header",
                "Row 1 is the merged group header \"G\" (no buttons); the buttons are on row 2, the column names.",
                "opens: no repair prompt",
                "autoFilter sheet=Report range=A2:B5",
                "merged sheet=Report cell=A1 mergeArea=A1:B1");

        // 3. each rolled-over sheet filters its own rows
        save("03-autofilter-rollover.xlsx", report().autoFilter(true).maxRows(4).write(items(10)));
        add("03-autofilter-rollover.xlsx", "autoFilter on rolled-over sheets",
                "Three sheets; each has buttons on row 1 and filters only its own rows.",
                "opens: no repair prompt",
                "sheets names=Report|Report (2)|Report (3)",
                "autoFilter sheet=Report range=A1:B5; autoFilter sheet=\"Report (2)\" range=A1:B5; autoFilter sheet=\"Report (3)\" range=A1:B3");

        // 4. a table brings its own filter buttons
        save("04-table-with-autofilter.xlsx", report().table("SampleTable").autoFilter(true).write(items(6)));
        add("04-table-with-autofilter.xlsx", "table with autoFilter(true)",
                "A table (banded rows) with filter buttons in its header. No repair prompt about an overlapping filter.",
                "opens: no repair prompt (this is the case that could be rejected)",
                "table sheet=Report name=SampleTable range=A1:B7 showAutoFilter=true");

        save("05-table-without-autofilter.xlsx", report().table("PlainTable").write(items(6)));
        add("05-table-without-autofilter.xlsx", "table without autoFilter",
                "The same table with NO filter buttons in the header.",
                "opens: no repair prompt",
                "table sheet=Report name=PlainTable range=A1:B7 showAutoFilter=false");

        // 5. chart per rolled-over sheet, below the summary rows
        save("06-chart-rollover-summary.xlsx", report().maxRows(4).summary(s -> s.sum("Qty").average("Qty"))
                .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).title("Qty").categoryColumn(0).valueColumn(1, "Qty"))
                .write(items(10)));
        add("06-chart-rollover-summary.xlsx", "one chart per sheet, below the summary",
                "Three sheets, each with a bar chart of its own rows that starts BELOW its Sum/Average rows and covers no cell with data.",
                "opens: no repair prompt",
                "chart sheet=Report count=1 topLeftRowGreaterThan=7; chart sheet=\"Report (2)\" count=1 topLeftRowGreaterThan=7; chart sheet=\"Report (3)\" count=1 topLeftRowGreaterThan=5",
                "visual: each chart plots the 4 / 4 / 2 items of its sheet");

        // 6. encryption
        save("07-encrypted-password-secret.xlsx", ExcelWriter.<Item>create().sheetName("Report").password("secret")
                .column("Name", Item::name).column("Qty", Item::qty).write(items(5)));
        add("07-encrypted-password-secret.xlsx", "encryption set on the writer",
                "Excel asks for a password; \"secret\" opens it, a wrong one is refused.",
                "opens with password=secret; opening with a wrong password must fail",
                "cell sheet=Report cell=A2 text=item 0");
        save("08-encrypted-late-password.xlsx", ExcelWriter.<Item>create().sheetName("Report")
                .column("Name", Item::name).column("Qty", Item::qty).write(items(5)), "Late-Binding1");
        add("08-encrypted-late-password.xlsx", "encryption given when the file is written",
                "Excel asks for a password; \"Late-Binding1\" opens it.",
                "opens with password=Late-Binding1; opening with a wrong password must fail",
                "cell sheet=Report cell=A2 text=item 0");

        // 7. formats and dates
        save("09-formats-and-dates.xlsx", ExcelWriter.<Object>create().sheetName("Report")
                .column("Date", x -> LocalDate.of(2024, 2, 29))
                .column("Early", x -> LocalDate.of(1900, 3, 1))
                .column("Moment", x -> LocalDateTime.of(2024, 12, 31, 23, 59, 58))
                .column("Millis", x -> LocalDateTime.of(2024, 12, 31, 23, 59, 58, 123_000_000), c -> c.type(ExcelDataType.DATETIME).format("yyyy-mm-dd hh:mm:ss.000"))
                .column("Integer", x -> 1234567, c -> c.type(ExcelDataType.INTEGER))
                .column("Double", x -> 1234.5, c -> c.type(ExcelDataType.DOUBLE))
                .column("Text number", x -> "0012")
                .column("Formula", x -> "E2*2", c -> c.type(ExcelDataType.FORMULA))
                .write(List.of(new Object())));
        add("09-formats-and-dates.xlsx", "number and date display",
                "Dates, a date after the 1900 leap-day quirk, a moment, a moment with milliseconds, grouped integers and decimals, a text that looks like a number, and a formula calculated when the file is opened.",
                "opens: no repair prompt",
                "cellText sheet=Report cell=A2 text=2024-02-29; cellText sheet=Report cell=B2 text=1900-03-01; cellText sheet=Report cell=C2 text=\"2024-12-31 23:59:58\"; cellText sheet=Report cell=D2 text=\"2024-12-31 23:59:58.123\"",
                "cellText sheet=Report cell=E2 text=1,234,567; cellText sheet=Report cell=F2 text=1,234.50; cellText sheet=Report cell=G2 text=0012 (text, not 12); cell sheet=Report cell=H2 value=2469134");

        // 8. protection
        save("10-protection.xlsx", ExcelWriter.<Item>create().sheetName("Report").protectSheet("lock").protectWorkbook("book")
                .column("Name", Item::name, c -> c.locked(false)).column("Qty", Item::qty, c -> c.locked(true)).write(items(4)));
        add("10-protection.xlsx", "sheet and workbook protection",
                "Name cells can be edited, Qty cells cannot; sheets cannot be added, renamed or deleted.",
                "opens: no repair prompt",
                "protection sheet=Report sheetProtected=true unlocked=A2,A3 locked=B2,B3; workbookStructureProtected=true");

        // 9. conditional formatting and validation
        save("11-conditional-and-validation.xlsx", ExcelWriter.<Item>create().sheetName("Report")
                .column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER).validation(ExcelValidation.integerBetween(0, 100)))
                .column("Status", x -> "Active", c -> c.dropdown("Active", "Inactive", "Pending"))
                .conditionalFormatting(cf -> cf.columns(1).greaterThan("3", ExcelColor.LIGHT_RED).dataBar(ExcelColor.BLUE))
                .write(items(8)));
        add("11-conditional-and-validation.xlsx", "conditional formatting and validation",
                "Qty above 3 is red with a data bar; Qty accepts only 0-100; Status is a dropdown of Active/Inactive/Pending.",
                "opens: no repair prompt",
                "validation sheet=Report cell=B2 type=whole; validation sheet=Report cell=C2 type=list",
                "conditionalFormats sheet=Report range=B2:B9 countAtLeast=2");

        // 10. pictures, links, rich text
        byte[] red = png(Color.RED);
        byte[] blue = png(Color.BLUE);
        // rich text keeps its formatting only in a workbook with shared strings
        save("12-images-links-richtext.xlsx", ExcelWriter.<Item>create(o -> o.useSharedStrings(true)).sheetName("Report")
                .column("Name", Item::name)
                .column("Photo", i -> ExcelImage.png(i.qty() % 2 == 0 ? red : blue), c -> c.type(ExcelDataType.IMAGE))
                .column("Link", i -> "https://example.com/" + i.name().replace(' ', '-'), c -> c.type(ExcelDataType.HYPERLINK))
                .column("Rich", i -> new ExcelRichText().text("Status: ").bold("APPROVED").text(" by ").styled("admin", s -> s.color(ExcelColor.BLUE).italic(true)),
                        c -> c.type(ExcelDataType.RICH_TEXT))
                .write(items(3)));
        add("12-images-links-richtext.xlsx", "pictures, hyperlinks and rich text",
                "A red or blue picture in every row, a working hyperlink, and rich text: \"APPROVED\" bold, \"admin\" blue and italic.",
                "opens: no repair prompt",
                "shapes sheet=Report pictures=3; hyperlinks sheet=Report count=3");

        // 11. several sheets
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("Users").tabColor(ExcelColor.GREEN).autoFilter().freezeRows(1)
                    .column("Name", Item::name).column("Qty", Item::qty).write(items(5).stream());
            book.<Item>sheet("Orders").tabColor(ExcelColor.BLUE).column("Name", Item::name).column("Qty", Item::qty)
                    .summary(s -> s.label("Total").sum("Qty")).write(items(4).stream());
            book.<Item>sheet("Archive").maxRows(2).table("ArchiveTable").autoFilter().column("Name", Item::name)
                    .write(items(3).stream());
            try (OutputStream out = Files.newOutputStream(dir.resolve("13-multi-sheet-workbook.xlsx"))) {
                book.finish().writeTo(out);
            }
        }
        add("13-multi-sheet-workbook.xlsx", "a workbook of several sheets",
                "Users (green tab, filter, first row frozen), Orders (blue tab, a Total row), Archive and Archive (2) (a table on each, with filter buttons).",
                "opens: no repair prompt",
                "sheets names=Users|Orders|Archive|Archive (2)",
                "autoFilter sheet=Users range=A1:B6; cell sheet=Orders cell=A6 text=Total; cell sheet=Orders cell=B6 value=6",
                "table sheet=Archive name=ArchiveTable_1 showAutoFilter=true; table sheet=\"Archive (2)\" name=ArchiveTable_2 showAutoFilter=true; freezePanes sheet=Users rows=1");

        // 12. layout: names, panes, print setup
        save("14-names-panes-print.xlsx", ExcelWriter.<Item>create().sheetName("Report").freezePane(1, 1).tabColor(255, 0, 0)
                .printSetup(ps -> ps.orientation(ExcelPrintSetup.Orientation.LANDSCAPE))
                .namedRange("Names", 0).column("Name", Item::name).column("Qty", Item::qty).write(items(5)));
        add("14-names-panes-print.xlsx", "named range, frozen panes, print setup",
                "A named range Names over A2:A6, the first row and column frozen, a red tab, landscape orientation.",
                "opens: no repair prompt",
                "name name=Names refersTo=Report!$A$2:$A$6; freezePanes sheet=Report rows=1 columns=1; pageSetup sheet=Report orientation=landscape");

        // 13. a filled-in template
        var buildTemplate = TemplatePreservationTest.class.getDeclaredMethod("buildTemplate");
        buildTemplate.setAccessible(true);
        byte[] template = (byte[]) buildTemplate.invoke(null);
        try (ExcelTemplateWriter writer = new ExcelTemplateWriter(new ByteArrayInputStream(template))) {
            writer.cell("B2", "A");
            writer.cell("A9", "Filled by the writer");
            writer.<Item>list(11).column("name", Item::name).column("qty", Item::qty).write(items(3));
            writer.sheet("Notes");
            writer.cell("A2", "second sheet cell");
            try (OutputStream out = Files.newOutputStream(dir.resolve("15-template-filled.xlsx"))) {
                writer.finish().writeTo(out);
            }
        }
        add("15-template-filled.xlsx", "a template, filled in",
                "The template's layout is intact: merged yellow title, a picture, a bar chart, the table SalesTable, a dropdown in B2, a hidden column F. The writer's cells are there too.",
                "opens: no repair prompt",
                "merged sheet=Report cell=A1 mergeArea=A1:C1; shapes sheet=Report pictures=1; chart sheet=Report count=1; table sheet=Report name=SalesTable range=A4:B7",
                "cell sheet=Report cell=B2 text=A; cell sheet=Report cell=A9 text=\"Filled by the writer\"; cell sheet=Report cell=A12 text=item 0; cell sheet=Notes cell=A2 text=\"second sheet cell\"");

        // 14. everything at once (no encryption): table + filter + chart + summary + group + rolled over + protected
        save("16-everything-at-once.xlsx", FeatureMatrixTest.buildWriter(FeatureMatrixTest.allFeaturesExceptEncryption())
                .write(IntStream.range(0, FeatureMatrixTest.ROWS).mapToObj(i -> new FeatureMatrixTest.Item("item " + i, i)).toList()));
        add("16-everything-at-once.xlsx", "every sheet feature combined",
                "Three sheets, each with a table (filter buttons), group header, summary rows, a chart, conditional formatting, protection, a row-number column and frozen header.",
                "opens: no repair prompt (the strictest case: all features together)",
                "sheets count=3");

        writeGapSample();
    }

    private void writeGapSample() throws Exception {
        // the summary starts after one empty row: does Excel then keep it out of the filter?
        save("17-autofilter-summary-after-gap.xlsx", report().autoFilter(true).summary(s -> s.sum("Qty").average("Qty"))
                .afterData(ctx -> ctx.getCurrentRow() + 1).write(items(7)));
        add("17-autofilter-summary-after-gap.xlsx", "autoFilter with a gap above the summary",
                "Like sample 01, but row 9 is empty and Sum/Average are on rows 10 and 11. Filter Name = \"item 3\": do Sum and Average stay visible now?",
                "opens: no repair prompt",
                "autoFilter sheet=Report range=A1:B8 (the range Excel shows right after opening, before any filter call)",
                "filterRows sheet=Report column=1 criteria=\"item 3\" expectVisibleDataRows=1; cell sheet=Report cell=A10 text=Sum must stay visible");
    }

    private void save(String file, ExcelHandler handler, String password) throws IOException {
        try (OutputStream out = Files.newOutputStream(dir.resolve(file))) {
            handler.writeTo(out, password);
        }
    }

    // ---- the descriptions --------------------------------------------------------------------------------

    private String manifest() {
        StringBuilder sb = new StringBuilder("{\n  \"files\": [\n");
        for (int i = 0; i < samples.size(); i++) {
            Sample s = samples.get(i);
            sb.append("    {\n      \"file\": ").append(q(s.file())).append(",\n      \"title\": ").append(q(s.title()))
                    .append(",\n      \"look\": ").append(q(s.look())).append(",\n      \"checks\": [");
            for (int c = 0; c < s.checks().size(); c++) sb.append(c == 0 ? "" : ", ").append(q(s.checks().get(c)));
            sb.append("]\n    }").append(i + 1 < samples.size() ? "," : "").append('\n');
        }
        return sb.append("  ]\n}\n").toString();
    }

    private String checklist() {
        StringBuilder sb = new StringBuilder("# Excel checks\n\nOpen each file in Excel. Every file must open **without** a message such as \"We found a problem with some content ... Do you want us to try to recover\".\n\n");
        for (Sample s : samples) {
            sb.append("## ").append(s.file()).append(" — ").append(s.title()).append("\n\n").append(s.look()).append("\n\n");
            for (String check : s.checks()) sb.append("- ").append(check).append('\n');
            sb.append('\n');
        }
        return sb.toString();
    }
}
