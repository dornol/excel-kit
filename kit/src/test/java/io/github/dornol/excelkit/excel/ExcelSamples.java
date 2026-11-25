package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.core.read.CellError;
import io.github.dornol.excelkit.core.read.RowError;
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
import java.util.LinkedHashMap;
import java.util.List;
import java.util.Map;
import java.util.stream.IntStream;
import java.util.stream.Stream;

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

    /** An always-true custom validation formula of about the given length: {@code AND(LEN(A2)>=0,...,TRUE)}. */
    private static String longValidationFormula(int length) {
        StringBuilder sb = new StringBuilder("AND(");
        while (sb.length() + "LEN(A2)>=0,TRUE)".length() <= length) sb.append("LEN(A2)>=0,");
        return sb.append("TRUE)").toString();
    }

    // ---- the samples -------------------------------------------------------------------------------------

    private void write() throws Exception {
        // 1. a filter over header and data, summary rows outside it
        save("01-autofilter-plain.xlsx", report().autoFilter(true).summary(s -> s.sum("Qty").average("Qty")).write(items(7)));
        add("01-autofilter-plain.xlsx", "autoFilter over the data",
                "Filter buttons on the header row (A1:B1). The file's filter is A1:B8, but Excel extends a filter over the rows that directly follow it, so Excel's own range is A1:B10 and Filter Name = \"item 3\" leaves ONE data row and hides Sum and Average too. This is the known Excel behaviour that sample 17 avoids; the checks below expect it.",
                "opens: no repair prompt",
                "headerVisible sheet=Report range=A1:B1",
                "autoFilter sheet=Report range=A1:B10 (Excel's range right after opening, before any filter call; the file itself holds A1:B8)",
                "filterRows sheet=Report column=1 criteria=\"item 3\" expectVisibleDataRows=1 (data range A2:A8) and rows 9 and 10 (Sum, Average) hidden: Excel's range includes them",
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
                "table sheet=Report name=SampleTable range=A1:B7 showAutoFilter=true",
                "headerVisible sheet=Report range=A1:B1 (a table header was white text on a white fill before)");

        save("05-table-without-autofilter.xlsx", report().table("PlainTable").write(items(6)));
        add("05-table-without-autofilter.xlsx", "table without autoFilter",
                "The same table with NO filter buttons in the header.",
                "opens: no repair prompt",
                "table sheet=Report name=PlainTable range=A1:B7 showAutoFilter=false",
                "headerVisible sheet=Report range=A1:B1");

        // 5. chart per rolled-over sheet, below the summary rows
        save("06-chart-rollover-summary.xlsx", report().maxRows(4).summary(s -> s.sum("Qty").average("Qty"))
                .chart(ch -> ch.type(ExcelChartConfig.ChartType.BAR).title("Qty").categoryColumn(0).valueColumn(1, "Qty").valueAxisMajorUnit(1))
                .write(items(10)));
        add("06-chart-rollover-summary.xlsx", "one chart per sheet, below the summary",
                "Three sheets, each with a bar chart of its own rows that starts BELOW its Sum/Average rows and covers no cell with data.",
                "opens: no repair prompt",
                "chart sheet=Report count=1 topLeftRowGreaterThan=7; chart sheet=\"Report (2)\" count=1 topLeftRowGreaterThan=7; chart sheet=\"Report (3)\" count=1 topLeftRowGreaterThan=5",
                "visual: each chart plots the 4 / 4 / 2 items of its sheet",
                "cellText sheet=\"Report (3)\" cell=B4 text=17 (Sum of 8 and 9, #,##0); cellText sheet=\"Report (3)\" cell=B5 text=8.5 (Average: 8.5, not 9, although the Qty cells are whole numbers; an average of 3 reads 3, not 3.00)",
                "chartAxis sheet=Report majorUnit=1 (valueAxisMajorUnit(1)); the value axis labels of the first chart read 0 1 2 3, not 0 1 1 2 2 3 3 4 (the Qty cells have the whole-number format #,##0 and Excel would choose 0.5)");

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
                .column("Text number", x -> "0012", c -> c.format(io.github.dornol.excelkit.excel.style.ExcelDataFormat.TEXT))
                .column("Formula", x -> "E2*2", c -> c.type(ExcelDataType.FORMULA))
                .write(List.of(new Object())));
        add("09-formats-and-dates.xlsx", "number and date display",
                "Dates, a date after the 1900 leap-day quirk, a moment, a moment with milliseconds, grouped integers and decimals, a text that looks like a number in the Text format (so that editing the cell does not turn 0012 into 12), and a formula calculated when the file is opened.",
                "opens: no repair prompt",
                "cellText sheet=Report cell=A2 text=2024-02-29; cellText sheet=Report cell=B2 text=1900-03-01; cellText sheet=Report cell=C2 text=\"2024-12-31 23:59:58\"; cellText sheet=Report cell=D2 text=\"2024-12-31 23:59:58.123\"",
                "cellText sheet=Report cell=E2 text=1,234,567; cellText sheet=Report cell=F2 text=1,234.50; cellText sheet=Report cell=G2 text=0012 (text, not 12); cellFormat sheet=Report cell=G2 numberFormat=@; edit sheet=Report cell=G2: enter the cell for editing (F2 or Range.Value2 = Range.Value2, then Enter) and report the value: it must still be the text 0012, not 12 (the Text format keeps it); cell sheet=Report cell=H2 value=2469134");

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
                "headerVisible sheet=Archive range=A1:A1; table sheet=Archive name=ArchiveTable_1 showAutoFilter=true; table sheet=\"Archive (2)\" name=ArchiveTable_2 showAutoFilter=true; freezePanes sheet=Users rows=1");

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
                "sheets count=3; headerVisible sheet=1 range=A2:C2 (the first sheet; row 2 holds the column names under the group header)");

        writeGapSample();
        writeFormSamples();
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

    // ---- upload forms: validations, hidden sheets, page breaks, merging, stripes, error reports --------------

    private record Choice(String status, String priority, String country) {
    }

    private record Pick(String category, String item, int code, String sub) {
    }

    private record Staff(String dept, String team, String name) {
    }

    private record Order(String dept, String item, int qty) {
    }

    private record Month(String name, int revenue, int target, double growth) {
    }

    private void writeFormSamples() throws Exception {
        // 18. the list validations: the in-cell arrow, the tooltip, blank cells
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Choice>sheet("Form")
                    .column("Status", Choice::status, c -> c.dropdown("Active", "Inactive", "Pending"))
                    .column("Priority", Choice::priority, c -> c.validation(ExcelValidation.list("High", "Low")
                            .allowBlank(false).inputTitle("Priority").inputMessage("Pick one of the two")))
                    .column("Country", Choice::country, c -> c.validation(ExcelValidation.listFromRange("Codes!$A$2:$A$4")))
                    .write(List.of(new Choice("Active", "High", "KR"), new Choice("Pending", "Low", "US")).stream());
            book.<String>sheet("Codes").column("Code", x -> x).write(Stream.of("KR", "US", "JP"));
            save("18-dropdown-validation.xlsx", book.finish());
        }
        add("18-dropdown-validation.xlsx", "dropdown validations",
                "Status (A), Priority (B) and Country (C) cells are dropdowns: click one and the list opens, WITH the arrow at the right of the cell. Selecting a Priority cell shows a tooltip titled \"Priority\" saying \"Pick one of the two\".",
                "opens: no repair prompt",
                "validation sheet=Form cell=A2 type=list inCellDropdown=true formula1=\"Active,Inactive,Pending\" (the arrow: this is the case that wrote showDropDown=true before, which hides it)",
                "validation sheet=Form cell=B2 type=list inCellDropdown=true ignoreBlank=false showInput=true inputTitle=Priority inputMessage=\"Pick one of the two\"",
                "validation sheet=Form cell=C2 type=list inCellDropdown=true ignoreBlank=true formula1==Codes!$A$2:$A$4",
                "visual: the arrow is visible at the right edge of A2, B2 and C2 when the cell is selected");

        // 19. dependent dropdowns, with a text parent and a numeric parent
        Map<String, List<String>> byCategory = new LinkedHashMap<>();
        byCategory.put("가", List.of("1", "2", "3"));
        byCategory.put("나", List.of("4", "5"));
        byCategory.put("다", List.of("6"));
        Map<String, List<String>> byCode = new LinkedHashMap<>();
        byCode.put("1", List.of("a", "b", "c"));
        byCode.put("2", List.of("d"));
        save("19-dependent-dropdown.xlsx", ExcelWriter.<Pick>create().sheetName("Form")
                .column("Category", Pick::category, c -> c.dropdown("가", "나", "다"))
                .column("Item", Pick::item, c -> c.dependentDropdown("Category", byCategory))
                .column("Code", Pick::code, c -> c.type(ExcelDataType.INTEGER))
                .column("Sub", Pick::sub, c -> c.dependentDropdown("Code", byCode))
                .write(List.of(new Pick("가", "1", 1, "a"), new Pick("나", "4", 2, "d"))));
        add("19-dependent-dropdown.xlsx", "dependent dropdowns",
                "Item (B) offers the options of the Category in the same row: 가 -> 1,2,3; 나 -> 4,5; 다 -> 6. Sub (D) offers the options of the NUMBER in Code (C): 1 -> a,b,c; 2 -> d. Two hidden helper sheets hold the options.",
                "opens: no repair prompt",
                "validation sheet=Form cell=B2 type=list inCellDropdown=true; validation sheet=Form cell=D2 type=list inCellDropdown=true",
                "dependentOptions sheet=Form cell=B2 parentCell=A2 parent=가 count=3; parent=나 count=2; parent=다 count=1; parent=zzz count=0 (an error such as #N/A also means an empty list)",
                "dependentOptions sheet=Form cell=D2 parentCell=C2 parent=1 (a NUMBER, not text) count=3; parent=2 count=1 (the numeric parent is the case a plain MATCH would miss)",
                "sheets names=Form|_excelkit_dep_1|_excelkit_dep_2; the two _excelkit_dep sheets are hidden (not very hidden)");

        // 20. hidden sheets and zoom
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("Codes").column("Name", Item::name).hideSheet().write(items(3).stream());
            book.<Item>sheet("Secret").column("Name", Item::name).veryHideSheet().write(items(3).stream());
            book.<Item>sheet("Main").zoom(90).column("Name", Item::name).column("Qty", Item::qty).write(items(5).stream());
            save("20-hidden-sheets-and-zoom.xlsx", book.finish());
        }
        add("20-hidden-sheets-and-zoom.xlsx", "hidden sheets and zoom",
                "Only the tab Main is shown, and the workbook opens on it (the first two sheets are hidden). Right-click a tab > Unhide lists Codes but NOT Secret (very hidden). Main opens at 90% zoom.",
                "opens: no repair prompt (a hidden first sheet must not be the one the workbook opens on)",
                "sheets names=Codes|Secret|Main (Worksheets lists hidden sheets too)",
                "sheetVisible sheet=Codes visible=hidden (xlSheetHidden=0); sheetVisible sheet=Secret visible=veryHidden (xlSheetVeryHidden=2); sheetVisible sheet=Main visible=visible (xlSheetVisible=-1)",
                "activeSheet name=Main (ActiveSheet.Name right after opening)",
                "zoom sheet=Main percent=90 (Activate the sheet, then ActiveWindow.Zoom)");

        // 21. page breaks
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("EveryThree").printSetup(p -> p.rowBreakEvery(3))
                    .column("Name", Item::name).column("Qty", Item::qty).write(items(8).stream());
            book.<Staff>sheet("ByDept").pageBreakOnChange(Staff::dept)
                    .column("Dept", Staff::dept).column("Name", Staff::name)
                    .write(List.of(new Staff("A", "", "1"), new Staff("A", "", "2"), new Staff("B", "", "3"),
                            new Staff("B", "", "4"), new Staff("B", "", "5"), new Staff("C", "", "6")).stream());
            save("21-page-breaks.xlsx", book.finish());
        }
        add("21-page-breaks.xlsx", "page breaks",
                "In Page Break Preview, EveryThree has a page break after every 3 data rows; ByDept has one where the department changes (A | B | C).",
                "opens: no repair prompt",
                "pageBreaks sheet=EveryThree manualRows=5,8 (the first row of each new page: HPageBreaks(i).Location.Row where Type=xlPageBreakManual=-4135; data rows are 2-9)",
                "pageBreaks sheet=ByDept manualRows=4,7 (A on rows 2-3, B on rows 4-6, C on row 7)",
                "HPageBreaks needs a printer driver; without one report \"not checkable\" and look at Page Break Preview instead");

        // 22. merging equal values
        save("22-merge-same-values.xlsx", ExcelWriter.<Staff>create().sheetName("Report")
                .column("Dept", Staff::dept, c -> c.mergeSameValues())
                .column("Team", Staff::team, c -> c.mergeSameValues())
                .column("Name", Staff::name)
                .write(List.of(new Staff("Sales", "East", "Ann"), new Staff("Sales", "East", "Bob"),
                        new Staff("Sales", "West", "Cy"), new Staff("Dev", "West", "Di"),
                        new Staff("Dev", "West", "Ed"), new Staff("Ops", "East", "Flo"))));
        add("22-merge-same-values.xlsx", "merged runs of equal values",
                "Dept shows Sales once over rows 2-4, Dev once over rows 5-6, Ops alone on row 7. Team merges only inside a department: East over rows 2-3, West over row 4 (single), West over rows 5-6 (NOT joined with row 4), East alone on row 7.",
                "opens: no repair prompt",
                "merged sheet=Report cell=A2 mergeArea=A2:A4; merged sheet=Report cell=A5 mergeArea=A5:A6; merged sheet=Report cell=A7 mergeArea=A7",
                "merged sheet=Report cell=B2 mergeArea=B2:B3; merged sheet=Report cell=B4 mergeArea=B4; merged sheet=Report cell=B5 mergeArea=B5:B6; merged sheet=Report cell=B7 mergeArea=B7",
                "info sheet=Report cell=A3: report what Value2 gives for the covered cell (the file keeps \"Sales\" there; Excel may show it as empty)",
                "visual: each merged value is vertically centered in its area");

        // 23. striped rows and the required mark
        save("23-striped-rows-required-mark.xlsx", ExcelWriter.<Item>create().sheetName("Report").headerColor(ExcelColor.STEEL_BLUE)
                .stripedRows()
                .column("Name", Item::name, c -> c.requiredMark())
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .write(items(6)));
        add("23-striped-rows-required-mark.xlsx", "striped rows and the required mark",
                "The header Name reads \"Name *\" (the star is a cell format; the cell value is still \"Name\"), Qty reads \"Qty\". Every second data row (rows 3, 5, 7) is light gray; rows 2, 4, 6 are not.",
                "opens: no repair prompt",
                "cellText sheet=Report cell=A1 text=\"Name *\"; cell sheet=Report cell=A1 value=Name; cellText sheet=Report cell=B1 text=Qty",
                "headerVisible sheet=Report range=A1:B1",
                "fill sheet=Report cells=A3,B3,A5,A7 rgb=F2F2F2; fill sheet=Report cells=A2,A4,A6 rgb=none (DisplayFormat.Interior.ColorIndex = xlColorIndexNone)");

        // 24. the error report of an upload
        List<RowError> errors = List.of(
                new RowError(2, 3, RowError.Type.MAPPING, List.of("age is not a number"), null,
                        List.of(new CellError(1, "Age", "abc", "not a number")), List.of("Bob", "abc")),
                new RowError(4, 5, RowError.Type.VALIDATION, List.of("Duplicate key [Name=Bob]; first seen at row 3"), null,
                        List.of(new CellError(0, "Name", "Bob", "duplicate of row 3")), List.of("Bob", "41")));
        save("24-error-report.xlsx", ExcelErrorReport.create(List.of("Name", "Age")).write(errors));
        add("24-error-report.xlsx", "upload error report",
                "Columns Row, Name, Age, Error. The failed cells are shaded light red and carry a comment: C2 (\"not a number\") and B3 (\"duplicate of row 3\"). The Error column is shaded too and shows the messages.",
                "opens: no repair prompt (cell comments written by the streaming writer)",
                "cell sheet=Errors cell=A2 value=3; cell sheet=Errors cell=A3 value=5; cell sheet=Errors cell=B2 text=Bob; cell sheet=Errors cell=C2 text=abc",
                "cellText sheet=Errors cell=D2 text=\"age is not a number\"; cellText sheet=Errors cell=D3 text=\"Duplicate key [Name=Bob]; first seen at row 3\"",
                "comment sheet=Errors cell=C2 text=\"not a number\"; comment sheet=Errors cell=B3 text=\"duplicate of row 3\"; comment sheet=Errors cell=B2 none",
                "fill sheet=Errors cells=C2,B3,D2,D3 rgb=FFC7CE; fill sheet=Errors cells=B2,C3 rgb=none");

        // 25. the required mark and stripes on a table
        save("25-required-mark-in-table.xlsx", ExcelWriter.<Item>create().sheetName("Report").table("Marked").autoFilter(true)
                .stripedRows()
                .column("Name", Item::name, c -> c.requiredMark())
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                .write(items(6)));
        add("25-required-mark-in-table.xlsx", "required mark and stripes on a table",
                "A table with filter buttons whose header Name reads \"Name *\" (a number format; the cell and the table column are still \"Name\"). Every second data row is light gray over the table style's banding.",
                "opens: no repair prompt (the header cell of a table carries a number format here)",
                "table sheet=Report name=Marked range=A1:B7 showAutoFilter=true column1=Name (ListColumns(1).Name must be the plain \"Name\")",
                "cellText sheet=Report cell=A1 text=\"Name *\"; cell sheet=Report cell=A1 value=Name; cellText sheet=Report cell=B1 text=Qty",
                "headerVisible sheet=Report range=A1:B1",
                "fill sheet=Report cells=A3,A5,A7 rgb=F2F2F2 (the stripe wins over the table style); info sheet=Report cell=A2: report the fill of a row that is not striped (the table's own banding)");

        // 26. features combined over a rolled-over sheet
        Map<String, List<String>> perDept = new LinkedHashMap<>();
        perDept.put("A", List.of("1", "2"));
        perDept.put("B", List.of("3"));
        perDept.put("C", List.of("4", "5", "6"));
        save("26-rollover-combined.xlsx", ExcelWriter.<Order>create().sheetName("Report").maxRows(4).autoFilter(true).zoom(80)
                .stripedRows().pageBreakOnChange(Order::dept)
                .column("Dept", Order::dept, c -> c.mergeSameValues())
                .column("Item", Order::item, c -> c.dependentDropdown("Dept", perDept))
                .column("Qty", Order::qty, c -> c.type(ExcelDataType.INTEGER))
                .write(List.of(new Order("A", "1", 1), new Order("A", "2", 2), new Order("A", "1", 3), new Order("B", "3", 4),
                        new Order("B", "3", 5), new Order("B", "3", 6), new Order("C", "4", 7), new Order("C", "5", 8),
                        new Order("C", "6", 9), new Order("C", "4", 10))));
        add("26-rollover-combined.xlsx", "merging, page breaks, dropdowns, stripes, filter and zoom over rolled-over sheets",
                "Three data sheets of 4, 4 and 2 rows and one hidden helper sheet shared by all of them. Dept is merged inside each sheet (a run does not continue onto the next sheet); a page starts where the department changes; Item is a dropdown that follows the Dept of its row; every second row is gray; each sheet is at 80% with its own filter.",
                "opens: no repair prompt (merged cells under a filter, on three sheets, with one helper sheet)",
                "sheets names=Report|Report (2)|Report (3)|_excelkit_dep_1; sheetVisible sheet=_excelkit_dep_1 visible=hidden",
                "merged sheet=Report cell=A2 mergeArea=A2:A4; merged sheet=Report cell=A5 mergeArea=A5; merged sheet=\"Report (2)\" cell=A2 mergeArea=A2:A3; merged sheet=\"Report (2)\" cell=A4 mergeArea=A4:A5; merged sheet=\"Report (3)\" cell=A2 mergeArea=A2:A3",
                "pageBreaks sheet=Report manualRows=5; pageBreaks sheet=\"Report (2)\" manualRows=4; pageBreaks sheet=\"Report (3)\" manualRows=(none)",
                "autoFilter sheet=Report range=A1:C5; autoFilter sheet=\"Report (2)\" range=A1:C5; autoFilter sheet=\"Report (3)\" range=A1:C3",
                "zoom sheet=Report percent=80; zoom sheet=\"Report (2)\" percent=80; zoom sheet=\"Report (3)\" percent=80",
                "dependentOptions sheet=Report cell=B2 parentCell=A2 parent=A count=2; parent=B count=1; parent=C count=3",
                "dependentOptions sheet=\"Report (3)\" cell=B2 parentCell=A2 parent=C count=3 (the helper sheet is shared)",
                "fill sheet=Report cells=C3,C5 rgb=F2F2F2; fill sheet=\"Report (2)\" cells=C3,C5 rgb=F2F2F2; fill sheet=\"Report (3)\" cells=C3 rgb=F2F2F2");

        // 27. group subtotals: SUBTOTAL formulas that a total below them must not count twice
        save("27-group-subtotal.xlsx", ExcelWriter.<Order>create().sheetName("Report").autoFilter(true)
                .column("Dept", Order::dept)
                .column("Item", Order::item)
                .column("Qty", Order::qty, c -> c.type(ExcelDataType.INTEGER))
                .groupSubtotal(Order::dept, s -> s.label("Item", k -> k + " total").sum("Qty"))
                .summary(s -> s.label("Dept", "Total").sum("Qty"))
                .write(List.of(new Order("A", "1", 10), new Order("A", "2", 20), new Order("B", "3", 5))));
        add("27-group-subtotal.xlsx", "group subtotals and a total that skips them",
                "Rows 2-3 are department A, row 4 is its subtotal (bold on gray), row 5 is department B, row 6 its subtotal, row 7 the total. The subtotals and the total are SUBTOTAL formulas: edit a Qty and they follow.",
                "opens: no repair prompt",
                "cell sheet=Report cell=C4 value=30; cell sheet=Report cell=C6 value=5",
                "cell sheet=Report cell=C7 value=35 (the total skips the two subtotal cells: 70 would mean they were counted twice)",
                "cellText sheet=Report cell=B4 text=\"A total\"; cellText sheet=Report cell=B6 text=\"B total\"",
                "fill sheet=Report cells=A4,B4,C4,A6,B6,C6 rgb=D9D9D9",
                "autoFilter sheet=Report range=A1:C7 (Excel's range right after opening; the file itself holds A1:C6, the subtotal rows without the total, and Excel extends a filter over the row directly below it)");

        // 28. subtotals over rolled-over sheets, with merging, stripes, page breaks and a filter
        save("28-subtotal-rollover.xlsx", ExcelWriter.<Order>create().sheetName("Report").maxRows(4).autoFilter(true)
                .stripedRows().pageBreakOnChange(Order::dept)
                .column("Dept", Order::dept, c -> c.mergeSameValues())
                .column("Item", Order::item)
                .column("Qty", Order::qty, c -> c.type(ExcelDataType.INTEGER))
                .groupSubtotal(Order::dept, s -> s.label("Item", k -> k + " total").sum("Qty"))
                .summary(s -> s.label("Item", "Total").sum("Qty"))
                .write(List.of(new Order("A", "1", 1), new Order("A", "2", 2), new Order("A", "1", 3), new Order("B", "3", 4),
                        new Order("B", "3", 5), new Order("B", "3", 6), new Order("C", "4", 7), new Order("C", "5", 8),
                        new Order("C", "6", 9), new Order("C", "4", 10))));
        add("28-subtotal-rollover.xlsx", "subtotals over rolled-over sheets with merging, stripes, page breaks and a filter",
                "Three sheets of 4, 4 and 2 data rows. Every sheet closes its last department with a subtotal and ends with its own total; a department that continues on the next sheet gets a second subtotal there. Dept is merged over the data rows of a run but not over the subtotal row; a page starts where the department changes (below the subtotal); data rows alternate white and light gray, subtotal rows are darker gray.",
                "opens: no repair prompt (merged cells, filter and subtotal formulas on three sheets)",
                "sheets names=Report|Report (2)|Report (3)",
                "cell sheet=Report cell=C5 value=6; cell sheet=Report cell=C7 value=4; cell sheet=Report cell=C8 value=10",
                "cell sheet=\"Report (2)\" cell=C4 value=11; cell sheet=\"Report (2)\" cell=C7 value=15; cell sheet=\"Report (2)\" cell=C8 value=26",
                "cell sheet=\"Report (3)\" cell=C4 value=19; cell sheet=\"Report (3)\" cell=C5 value=19",
                "merged sheet=Report cell=A2 mergeArea=A2:A4; merged sheet=Report cell=A5 mergeArea=A5; merged sheet=\"Report (2)\" cell=A2 mergeArea=A2:A3; merged sheet=\"Report (2)\" cell=A5 mergeArea=A5:A6; merged sheet=\"Report (3)\" cell=A2 mergeArea=A2:A3",
                "pageBreaks sheet=Report manualRows=6; pageBreaks sheet=\"Report (2)\" manualRows=5; pageBreaks sheet=\"Report (3)\" manualRows=(none)",
                "autoFilter sheet=Report range=A1:C8; autoFilter sheet=\"Report (2)\" range=A1:C8; autoFilter sheet=\"Report (3)\" range=A1:C5 (Excel's range right after opening; the files hold A1:C7, A1:C7 and A1:C4, and Excel extends each over the total directly below)",
                "fill sheet=Report cells=C3,C6 rgb=F2F2F2; fill sheet=Report cells=C5,C7 rgb=D9D9D9; fill sheet=\"Report (2)\" cells=C3,C6 rgb=F2F2F2; fill sheet=\"Report (2)\" cells=C4,C7 rgb=D9D9D9");

        // 29. subtotals on a protected sheet whose data columns are unlocked
        save("29-subtotal-protection.xlsx", ExcelWriter.<Order>create().sheetName("Report").protectSheet("pw")
                .column("Dept", Order::dept)
                .column("Item", Order::item, c -> c.locked(false))
                .column("Qty", Order::qty, c -> c.type(ExcelDataType.INTEGER).locked(false))
                .groupSubtotal(Order::dept, s -> s.label("Item", k -> k + " total").sum("Qty"))
                .write(List.of(new Order("A", "1", 10), new Order("A", "2", 20), new Order("B", "3", 5))));
        add("29-subtotal-protection.xlsx", "subtotals stay locked on a protected sheet",
                "The sheet is protected (password pw). Item and Qty are unlocked for the data rows, so a person can type there; the subtotal rows (4 and 6) are locked in every column, so a subtotal cannot be overwritten.",
                "opens: no repair prompt",
                "protection sheet=Report ProtectContents=True",
                "protection sheet=Report cell=B2 Locked=False; cell=C2 Locked=False; cell=A2 Locked=True",
                "protection sheet=Report cell=B4 Locked=True; cell=C4 Locked=True; cell=A4 Locked=True; cell=C6 Locked=True",
                "edit sheet=Report cell=C2 value=20 must succeed, then cell sheet=Report cell=C4 value=40 (the subtotal follows the edit); restore C2 to 10",
                "edit sheet=Report cell=C4 value=1 must FAIL (the cell is locked): report the exception text");

        // 30. subtotals under a filter, a conditional format and validations
        save("30-subtotal-filter-and-rules.xlsx", ExcelWriter.<Order>create().sheetName("Report").autoFilter(true)
                .column("Dept", Order::dept, c -> c.dropdown("A", "B"))
                .column("Item", Order::item)
                .column("Qty", Order::qty, c -> c.type(ExcelDataType.INTEGER).validation(ExcelValidation.integerBetween(0, 100)))
                .conditionalFormatting(c -> c.columns(2).greaterThan("15", ExcelColor.LIGHT_YELLOW))
                .groupSubtotal(Order::dept, s -> s.label("Item", k -> k + " total").keyColumn("Dept").sum("Qty"))
                .summary(s -> s.label("Item", "Total").sum("Qty"))
                .write(List.of(new Order("A", "1", 10), new Order("A", "2", 20), new Order("B", "3", 5), new Order("B", "4", 30))));
        add("30-subtotal-filter-and-rules.xlsx", "a total under a filter, a conditional format that skips the subtotal rows, and validations over them",
                "Rows 2-3 are A, row 4 its subtotal (Dept \"A\", Item \"A total\": keyColumn shows the key), rows 5-6 are B, row 7 its subtotal, row 8 the total. Qty above 15 is yellow in the data rows (C3 = 20 and C6 = 30) but NOT in the subtotal rows, although both subtotals (30 and 35) are above 15: the rule covers C2:C3 and C5:C6 only, so the subtotal rows keep their bold gray. Dept is a dropdown (A, B) and Qty accepts whole numbers 0-100, down through the subtotal rows.",
                "opens: no repair prompt",
                "cell sheet=Report cell=C4 value=30; cell sheet=Report cell=C7 value=35; cell sheet=Report cell=C8 value=65",
                "cellFormat sheet=Report cell=C8 numberFormat=#,##0 fontBold=True fill=gray (the total row is styled like the subtotal rows and takes the format of its column); cellFormat sheet=Report cell=A8 fontBold=True fill=gray (the band runs across the whole row)",
                "cellText sheet=Report cell=A4 text=A; cellText sheet=Report cell=B4 text=\"A total\"; cellText sheet=Report cell=A7 text=B",
                "filterRows sheet=Report column=1 criteria=\"A\" expectVisibleDataRows=3 (rows 2, 3 and the subtotal row 4, which holds the key A in Dept: an exact filter keeps it); then cell sheet=Report cell=C8 value=30 (SUBTOTAL skips the rows the filter hides and the visible subtotal cell); ShowAllData and cell C8 value=65 again; info sheet=Report cell=A8: report whether row 8 (the total, empty in Dept) is hidden by the filter",
                "conditionalFormats sheet=Report one rule over C2:C3 and C5:C6 (the data rows; not C4, C7 or C8); fill sheet=Report cells=C3,C6 rgb=FFEB9C (the rule's color, see info); fill sheet=Report cells=C4,C7 must show the gray of the subtotal style, not the rule's yellow",
                "validation sheet=Report cell=A2 type=3; validation sheet=Report cell=C2 type=1; info sheet=Report cell=C4: report Validation.Type of the subtotal cell (the validation runs below the header, through the subtotal rows)");

        // 31. a combination chart: columns, a line on the same axis and a line on a second axis
        save("31-combo-chart.xlsx", ExcelWriter.<Month>create().sheetName("Sales")
                .column("Month", Month::name)
                .column("Revenue", Month::revenue, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Target", Month::target, c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
                .column("Growth", Month::growth, c -> c.type(ExcelDataType.DOUBLE_PERCENT))
                .chart(c -> c.type(ExcelChartConfig.ChartType.BAR).title("Revenue, target and growth")
                        .categoryColumn(0).valueColumn(1, "Revenue").lineColumn(2, "Target").secondaryLineColumn(3, "Growth")
                        .valueAxisTitle("Revenue").secondaryValueAxisTitle("Growth")
                        .legendPosition(ExcelChartConfig.LegendPosition.BOTTOM).position(5, 1, 14, 18))
                .write(List.of(new Month("Jan", 100, 110, 0.05), new Month("Feb", 120, 110, 0.20), new Month("Mar", 90, 120, -0.25),
                        new Month("Apr", 150, 130, 0.67), new Month("May", 170, 150, 0.13), new Month("Jun", 160, 170, -0.06))));
        add("31-combo-chart.xlsx", "a combination chart: columns, a line on the same axis, a line on a second axis",
                "Six months. Columns of Revenue, a line of Target over them on the same left axis (\"Revenue\"), and a line of Growth (percent) on its own axis on the right (\"Growth\"), which is a different scale: the growth line must not sit flat at the bottom of the revenue scale. The legend is below; the chart is right of the data.",
                "opens: no repair prompt (three chart groups on two pairs of axes, one category axis hidden)",
                "chart sheet=Sales count=1",
                "chartSeries sheet=Sales index=1 chartType=xlColumnClustered(51) axisGroup=1; chartSeries sheet=Sales index=2 chartType=xlLine or xlLineMarkers axisGroup=1; chartSeries sheet=Sales index=3 chartType=xlLine or xlLineMarkers axisGroup=2 (the secondary axis)",
                "chartAxis sheet=Sales hasSecondaryValueAxis=True; chartAxis sheet=Sales hasSecondaryCategoryAxis=False (it is stored but not shown)",
                "chartAxisTitle sheet=Sales primary=Revenue secondary=Growth",
                "visual: export the chart to a PNG (Chart.Export) and look: three series, a left and a right value axis");

        // 32. an empty sheet that says so, next to a sheet with rows
        try (ExcelWorkbook book = ExcelWorkbook.create()) {
            book.<Item>sheet("Empty").autoFilter().freezeRows(1).emptyMessage("No data")
                    .column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                    .column("Note", i -> "")
                    .summary(s -> s.label("Name", "Total").sum("Qty"))
                    .write(List.of());
            book.<Item>sheet("Rows").autoFilter().emptyMessage("No data")
                    .column("Name", Item::name).column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER))
                    .column("Note", i -> "")
                    .write(items(3));
            save("32-empty-message.xlsx", book.finish());
        }
        add("32-empty-message.xlsx", "an empty sheet says so; a sheet with rows is untouched",
                "Sheet Empty: the header, and in row 2 \"No data\" centered over A:C in gray italics (one merged cell); below it a summary row \"Total\" with no formulas (nothing to add up). Sheet Rows asked for the same message but has rows: the header and three rows, no message, no merged cell.",
                "opens: no repair prompt (a merged line under a filter, with a summary below it)",
                "merged sheet=Empty cell=A2 mergeArea=A2:C2; merged sheet=Rows cell=A2 mergeArea=A2",
                "cellText sheet=Empty cell=A2 text=\"No data\"; cellText sheet=Empty cell=A3 text=Total; cellText sheet=Rows cell=A2 text=\"item 0\"",
                "cellFormat sheet=Empty cell=A2 horizontalAlignment=xlCenter(-4108) fontItalic=True fontColor=808080",
                "autoFilter sheet=Empty range=A1:C1 (the header only; Excel may extend it over the rows directly below it, so report the range it shows); freezePanes sheet=Empty rows=1",
                "cell sheet=Empty cell=B3 value=(empty): no SUM over the line of text");

        // 33. custom validation formulas around Excel's limit of 255 characters for a validation formula
        String shortFormula = longValidationFormula(240);
        String longFormula = longValidationFormula(300);
        save("33-long-validation-formula.xlsx", ExcelWriter.<Item>create().sheetName("Report")
                .column("Name", Item::name, c -> c.validation(ExcelValidation.formula(shortFormula)))
                .column("Qty", Item::qty, c -> c.type(ExcelDataType.INTEGER).validation(ExcelValidation.formula(longFormula)))
                .write(items(3)));
        add("33-long-validation-formula.xlsx", "custom validation formulas of " + shortFormula.length() + " and "
                        + longFormula.length() + " characters",
                "Sheet Report: a header and three rows. Column A holds a custom validation whose formula is "
                        + shortFormula.length() + " characters, column B one of " + longFormula.length()
                        + " characters (Excel documents 255 as the limit of a validation formula). Both formulas are always true "
                        + "(AND of LEN(...)>=0 terms), so any value is accepted. This file decides whether the library has to refuse "
                        + "a validation formula over 255 characters: report exactly what Excel does with the second one.",
                "opens: no repair prompt for column A (under the limit); report whether column B opens clean, is repaired, or has its validation dropped",
                "validation sheet=Report cell=A2 type=custom formula1.length=" + shortFormula.length(),
                "validation sheet=Report cell=B2 type=custom formula1.length=" + longFormula.length()
                        + " (report whether the validation is still there after Excel opens the file, and its formula1 length)",
                "typing: type any text into Report!A5 and Report!B5 (Range.Value2): both must be accepted");
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
