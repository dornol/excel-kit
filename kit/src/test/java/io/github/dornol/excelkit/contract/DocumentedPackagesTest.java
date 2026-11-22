package io.github.dornol.excelkit.contract;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Keeps the "Packages and Imports" tables in the reference guide and in {@code AI.md} equal to the public types
 * of the source tree, so a type that is added, renamed or moved cannot leave a wrong package in the docs.
 */
class DocumentedPackagesTest {

    private static final String BASE = "io.github.dornol.excelkit";
    private static final List<Path> SOURCE_ROOTS = List.of(
            Path.of("src/main/java"), Path.of("../excel-kit-spring/src/main/java"));
    private static final List<Path> DOCS = List.of(
            Path.of("../docs/guide/reference.md"), Path.of("src/main/resources/META-INF/AI.md"));

    private static final Pattern PACKAGE = Pattern.compile("^package ([\\w.]+);", Pattern.MULTILINE);
    private static final Pattern ROW = Pattern.compile("^\\| `(" + Pattern.quote(BASE) + "[\\w.]*)` \\| (.*?) \\|", Pattern.MULTILINE);
    private static final Pattern NAME = Pattern.compile("`(\\w+)`");

    @Test
    void everyPublicTypeIsListedUnderItsPackage() throws IOException {
        assumeTrue(SOURCE_ROOTS.stream().allMatch(Files::isDirectory), "sources of both modules are needed");
        Map<String, Set<String>> actual = publicTypesByPackage();
        List<String> problems = new ArrayList<>();
        for (Path doc : DOCS) {
            assumeTrue(Files.isRegularFile(doc), doc + " is not available");
            Map<String, Set<String>> documented = documentedTypes(Files.readString(doc));
            for (var entry : actual.entrySet()) {
                Set<String> listed = documented.getOrDefault(entry.getKey(), Set.of());
                Set<String> missing = new TreeSet<>(entry.getValue());
                missing.removeAll(listed);
                if (!missing.isEmpty()) {
                    problems.add(doc + ": package " + entry.getKey() + " is missing " + missing);
                }
                Set<String> stale = new TreeSet<>(listed);
                stale.removeAll(entry.getValue());
                if (!stale.isEmpty()) {
                    problems.add(doc + ": package " + entry.getKey() + " lists types it does not have: " + stale);
                }
            }
            Set<String> unknown = new TreeSet<>(documented.keySet());
            unknown.removeAll(actual.keySet());
            if (!unknown.isEmpty()) {
                problems.add(doc + ": lists packages without public types: " + unknown);
            }
        }
        assertTrue(problems.isEmpty(), "Documented packages differ from the source:\n" + String.join("\n", problems));
    }

    private static Map<String, Set<String>> publicTypesByPackage() throws IOException {
        Map<String, Set<String>> result = new TreeMap<>();
        for (Path root : SOURCE_ROOTS) {
            try (Stream<Path> files = Files.walk(root)) {
                for (Path file : (Iterable<Path>) files.filter(f -> f.toString().endsWith(".java")
                        && !f.getFileName().toString().equals("package-info.java"))::iterator) {
                    String source = Files.readString(file);
                    String type = file.getFileName().toString().replace(".java", "");
                    Matcher pkg = PACKAGE.matcher(source);
                    if (pkg.find() && Pattern.compile("^public (?:(?:final|abstract|sealed|non-sealed|static) )*"
                            + "(?:class|interface|enum|record|@interface) " + type + "\\b", Pattern.MULTILINE)
                            .matcher(source).find()) {
                        result.computeIfAbsent(pkg.group(1), k -> new TreeSet<>()).add(type);
                    }
                }
            }
        }
        return result;
    }

    private static Map<String, Set<String>> documentedTypes(String doc) {
        Map<String, Set<String>> result = new TreeMap<>();
        Matcher row = ROW.matcher(doc);
        while (row.find()) {
            Set<String> names = new TreeSet<>();
            Matcher name = NAME.matcher(row.group(2));
            while (name.find()) {
                names.add(name.group(1));
            }
            result.put(row.group(1), names);
        }
        return result;
    }
}
