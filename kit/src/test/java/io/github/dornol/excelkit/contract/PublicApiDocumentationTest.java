package io.github.dornol.excelkit.contract;

import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.ArrayList;
import java.util.List;
import java.util.regex.Matcher;
import java.util.regex.Pattern;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.junit.jupiter.api.Assumptions.assumeTrue;

/**
 * Keeps the public API documented: every public top-level type says in its class javadoc which release added it
 * ({@code @since 0.21.0}, not a date), and every public method or constructor of such a type has a javadoc. Overrides
 * ({@code @Override}) are exempt; they inherit the contract.
 */
class PublicApiDocumentationTest {

    private static final List<Path> SOURCE_ROOTS = List.of(
            Path.of("src/main/java"), Path.of("../excel-kit-spring/src/main/java"));

    private static final Pattern PUBLIC_TYPE = Pattern.compile(
            "^public (?:(?:final|abstract|sealed|non-sealed|static) )*(?:class|interface|enum|record|@interface) (\\w+)",
            Pattern.MULTILINE);
    private static final Pattern PUBLIC_MEMBER = Pattern.compile(
            "^    public (?!(?:static |final |abstract )*(?:class|enum|record|interface)\\b)[^=;]*\\(.*");
    private static final Pattern SINCE_VERSION = Pattern.compile("@since \\d+\\.\\d+");

    @Test
    void publicTypesAndMethodsAreDocumented() throws IOException {
        assumeTrue(SOURCE_ROOTS.stream().allMatch(Files::isDirectory), "sources of both modules are needed");
        List<String> problems = new ArrayList<>();
        for (Path root : SOURCE_ROOTS) {
            try (Stream<Path> files = Files.walk(root)) {
                for (Path file : (Iterable<Path>) files.filter(f -> f.toString().endsWith(".java")
                        && !f.getFileName().toString().equals("package-info.java"))::iterator) {
                    check(root, file, problems);
                }
            }
        }
        assertTrue(problems.isEmpty(), "Public API documentation is incomplete:\n" + String.join("\n", problems));
    }

    private static void check(Path root, Path file, List<String> problems) throws IOException {
        String source = Files.readString(file);
        String name = file.getFileName().toString().replace(".java", "");
        Matcher type = PUBLIC_TYPE.matcher(source);
        if (!type.find() || !type.group(1).equals(name)) {
            return;
        }
        String where = root.relativize(file).toString();
        String[] lines = source.split("\n", -1);
        int declaration = lineOf(source, type.start());
        String classDoc = docAbove(lines, declaration);
        if (classDoc == null) {
            problems.add(where + ": " + name + " has no class javadoc");
        } else if (!SINCE_VERSION.matcher(classDoc).find()) {
            problems.add(where + ": " + name + " has no version-style @since in its class javadoc");
        }
        for (int i = declaration + 1; i < lines.length; i++) {
            if (PUBLIC_MEMBER.matcher(lines[i]).matches() && !isOverride(lines, i) && docAbove(lines, i) == null) {
                problems.add(where + ":" + (i + 1) + ": no javadoc on " + lines[i].trim());
            }
        }
    }

    private static int lineOf(String source, int offset) {
        int line = 0;
        for (int i = 0; i < offset; i++) {
            if (source.charAt(i) == '\n') {
                line++;
            }
        }
        return line;
    }

    /** The javadoc directly above {@code line}, skipping annotations, or {@code null} when there is none. */
    private static String docAbove(String[] lines, int line) {
        int i = line - 1;
        while (i >= 0 && lines[i].trim().startsWith("@")) {
            i--;
        }
        if (i < 0 || !lines[i].trim().endsWith("*/")) {
            return null;
        }
        int end = i;
        while (i >= 0 && !lines[i].contains("/**")) {
            i--;
        }
        return i < 0 ? null : String.join("\n", List.of(lines).subList(i, end + 1));
    }

    private static boolean isOverride(String[] lines, int line) {
        for (int i = line - 1; i >= 0 && lines[i].trim().startsWith("@"); i--) {
            if (lines[i].trim().startsWith("@Override")) {
                return true;
            }
        }
        return false;
    }
}
