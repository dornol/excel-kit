package io.github.dornol.excelkit.csv;

import io.github.dornol.excelkit.core.ReadResult;
import org.junit.jupiter.api.Test;
import org.junit.jupiter.api.io.TempDir;

import java.io.IOException;
import java.nio.charset.Charset;
import java.nio.charset.StandardCharsets;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * What a {@link CsvWriter} writes, a {@link CsvReader} given the same delimiter and charset reads back: every
 * combination of delimiter, charset, BOM, quoting, injection defense and line separator is written with
 * values that are hard for CSV (separators, quotes, line breaks, padding, formula characters, non-ASCII) and
 * read back. The only expected differences are the documented ones: {@code null} comes back empty, and the
 * injection defense puts a {@code '} in front of text that could be taken for a formula.
 */
class CsvRoundTripCombinationTest {

    @TempDir
    Path tempDir;

    private record Item(String name, Number score) {
    }

    private static final List<Item> ITEMS = List.of(
            new Item("plain", 1),
            new Item("a,b", -2),
            new Item("say \"hi\"", 3.5),
            new Item("line1\nline2", 4),
            new Item("crlf\r\nx", 5),
            new Item(" padded ", 6),
            new Item("", 7),
            new Item(null, 8),
            new Item("=1+1", -9),
            new Item("+x", 10),
            new Item("-y", 11),
            new Item("@z", 12),
            new Item("  =lead", 13),
            new Item("\uD55C\uAE00", 14),
            new Item("emoji \uD83D\uDE00", 15),
            new Item("tab\there", 16),
            new Item("semi;colon", 17),
            new Item("pipe|bar", 18),
            new Item("trailing\\", 19),
            new Item("'quote", 20),
            new Item("\"", 21),
            new Item("\t=tabbed", 22));

    private static final char[] DELIMITERS = {',', ';', '\t', '|'};
    private static final List<Charset> CHARSETS = List.of(StandardCharsets.UTF_8, StandardCharsets.UTF_16,
            Charset.forName("windows-1252"), Charset.forName("EUC-KR"));
    private static final String[] LINE_SEPARATORS = {"\r\n", "\n", "\r"};

    @Test
    void everyCombinationOfWriterOptionsReadsBack() throws IOException {
        List<String> problems = new ArrayList<>();
        int combinations = 0;
        for (char delimiter : DELIMITERS) {
            for (Charset charset : CHARSETS) {
                for (boolean bom : new boolean[]{true, false}) {
                    for (CsvQuoting quoting : CsvQuoting.values()) {
                        for (boolean defense : new boolean[]{true, false}) {
                            for (String lineSeparator : LINE_SEPARATORS) {
                                combinations++;
                                roundTrip(problems, delimiter, charset, bom, quoting, defense, lineSeparator);
                            }
                        }
                    }
                }
            }
        }
        assertEquals(4 * 4 * 2 * 3 * 2 * 3, combinations);
        assertTrue(problems.isEmpty(), problems.size() + " round trip problem(s):\n  " + String.join("\n  ", problems));
    }

    @Test
    void everyDialectReadsBack() throws IOException {
        List<String> problems = new ArrayList<>();
        for (CsvDialect dialect : CsvDialect.values()) {
            Path file = tempDir.resolve(dialect + ".csv");
            CsvWriter.<Item>create().dialect(dialect).column("Name", Item::name).column("Score", Item::score)
                    .write(ITEMS).writeTo(file);
            compare(problems, "dialect " + dialect, file, CsvReader.forMap().dialect(dialect), ITEMS, true);
        }
        assertTrue(problems.isEmpty(), problems.size() + " round trip problem(s):\n  " + String.join("\n  ", problems));
    }

    private void roundTrip(List<String> problems, char delimiter, Charset charset, boolean bom, CsvQuoting quoting,
                           boolean defense, String lineSeparator) throws IOException {
        String label = "delimiter=" + (delimiter == '\t' ? "TAB" : String.valueOf(delimiter)) + " charset=" + charset
                + " bom=" + bom + " quoting=" + quoting + " defense=" + defense + " lineSeparator="
                + lineSeparator.replace("\r", "CR").replace("\n", "LF");
        // A charset that cannot hold a value cannot round trip it; that failure is tested on its own.
        List<Item> items = ITEMS.stream().filter(item -> item.name() == null || charset.newEncoder().canEncode(item.name())).toList();
        Path file = tempDir.resolve("round-trip-" + Math.abs(label.hashCode()) + ".csv");
        try {
            CsvWriter.<Item>create().delimiter(delimiter).charset(charset).bom(bom).quoting(quoting)
                    .csvInjectionDefense(defense).lineSeparator(lineSeparator)
                    .column("Name", Item::name).column("Score", Item::score)
                    .write(items).writeTo(file);
        } catch (RuntimeException e) {
            problems.add("[" + label + "] failed to write: " + e);
            return;
        }
        compare(problems, label, file, CsvReader.forMap().delimiter(delimiter).charset(charset), items, defense);
    }

    /** The value the reader should see for what the writer was given. */
    private static String expected(Item item, boolean defense) {
        String name = item.name() == null ? "" : item.name();
        if (defense && !name.isEmpty()) {
            int first = 0;
            while (first < name.length() && name.charAt(first) == ' ') first++;
            if ("=+-@\t\r".indexOf(name.charAt(0)) >= 0 || (first < name.length() && "=+-@\t\r".indexOf(name.charAt(first)) >= 0)) {
                return "'" + name;
            }
        }
        return name;
    }

    private static void compare(List<String> problems, String label, Path file, CsvReader<Map<String, String>> reader,
                                List<Item> items, boolean defense) {
        List<ReadResult<Map<String, String>>> results = new ArrayList<>();
        try {
            reader.read(file, results::add);
        } catch (RuntimeException e) {
            problems.add("[" + label + "] failed to read: " + e + (e.getCause() == null ? "" : " / " + e.getCause()));
            return;
        }
        if (results.size() != items.size()) {
            problems.add("[" + label + "] " + items.size() + " rows written, " + results.size() + " read");
            return;
        }
        for (int i = 0; i < items.size(); i++) {
            ReadResult<Map<String, String>> result = results.get(i);
            Item item = items.get(i);
            if (!result.success()) {
                problems.add("[" + label + "] row " + i + " did not read: " + result.messages());
                continue;
            }
            String name = result.data().get("Name");
            String score = result.data().get("Score");
            if (!expected(item, defense).equals(name == null ? "" : name)) {
                problems.add("[" + label + "] Name of row " + i + ": wrote <" + show(item.name()) + ">, expected <"
                        + show(expected(item, defense)) + ">, read <" + show(name) + ">");
            }
            if (!String.valueOf(item.score()).equals(score)) {
                problems.add("[" + label + "] Score of row " + i + ": wrote <" + item.score() + ">, read <" + show(score) + ">");
            }
        }
    }

    private static String show(String value) {
        return value == null ? "null" : value.replace("\r", "\\r").replace("\n", "\\n").replace("\t", "\\t");
    }
}
