package io.github.dornol.excelkit.testsupport;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.regex.Pattern;
import java.util.stream.Stream;

/**
 * Counts the temp files the library can leave under {@code java.io.tmpdir}: Apache POI sheet files
 * in {@code poifiles}, and the UUID-named directories that CSV output and reader input are staged in.
 * JUnit's own {@code @TempDir} folders do not match the UUID pattern.
 */
public final class TempFiles {
    private static final Pattern UUID_NAMED =
            Pattern.compile("[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}.*");

    private TempFiles() {
    }

    public static long leftoverCount() throws IOException {
        Path tmp = Path.of(System.getProperty("java.io.tmpdir"));
        long count = 0;
        Path poi = tmp.resolve("poifiles");
        if (Files.isDirectory(poi)) {
            try (Stream<Path> files = Files.list(poi)) {
                count += files.count();
            }
        }
        try (Stream<Path> entries = Files.list(tmp)) {
            count += entries.filter(path -> UUID_NAMED.matcher(path.getFileName().toString()).matches()).count();
        }
        return count;
    }
}
