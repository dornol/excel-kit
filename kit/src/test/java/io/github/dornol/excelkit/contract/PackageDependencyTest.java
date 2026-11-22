package io.github.dornol.excelkit.contract;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Pins the layering of the library's packages by reading the main sources, so a stray import cannot bring
 * back a package cycle (core depending on excel and csv, excel and csv depending on each other, ...).
 *
 * <pre>
 * core  &lt;-  core.io  &lt;-  core.detect, core.read  &lt;-  csv, excel  &lt;-  schema
 * </pre>
 */
class PackageDependencyTest {

    private static final String BASE = "io.github.dornol.excelkit.";
    private static final Path SOURCES = Path.of("src/main/java/io/github/dornol/excelkit");

    /** Package (relative to {@code io.github.dornol.excelkit}) to the other packages it may reference. */
    private static final Map<String, Set<String>> ALLOWED = Map.of(
            "core", Set.of(),
            "core.io", Set.of("core"),
            "core.detect", Set.of("core", "core.io"),
            "core.read", Set.of("core", "core.io"),
            "csv", Set.of("core", "core.io", "core.detect", "core.read"),
            "excel", Set.of("core", "core.io", "core.detect", "core.read"),
            "schema", Set.of("core", "core.io", "core.detect", "core.read", "csv", "excel")
    );

    private static final Pattern REFERENCE = Pattern.compile(
            "\\bio\\.github\\.dornol\\.excelkit\\.(core\\.(?:io|detect|read)|core|csv|excel|schema)\\.[A-Za-z]");

    @Test
    void packagesOnlyDependOnLowerLayers() throws IOException {
        List<String> violations = new ArrayList<>();
        try (Stream<Path> files = Files.walk(SOURCES)) {
            for (Path file : (Iterable<Path>) files.filter(f -> f.toString().endsWith(".java"))::iterator) {
                String pkg = packageOf(file);
                Set<String> allowed = ALLOWED.get(pkg);
                assertTrue(allowed != null, "Package " + pkg + " has no entry in ALLOWED; add its layer: " + file);
                Set<String> used = new TreeSet<>();
                Matcher m = REFERENCE.matcher(withoutComments(Files.readString(file)));
                while (m.find()) {
                    used.add(m.group(1));
                }
                used.remove(pkg);
                used.removeAll(allowed);
                if (!used.isEmpty()) {
                    violations.add(SOURCES.relativize(file) + " must not depend on " + used);
                }
            }
        }
        assertTrue(violations.isEmpty(), "Package layering violated:\n" + String.join("\n", violations));
    }

    private static String packageOf(Path file) {
        Path dir = SOURCES.relativize(file).getParent();
        return dir.toString().replace('/', '.').replace('\\', '.');
    }

    private static String withoutComments(String source) {
        return source.replaceAll("(?s)/\\*.*?\\*/", "").replaceAll("//[^\\n]*", "");
    }
}
