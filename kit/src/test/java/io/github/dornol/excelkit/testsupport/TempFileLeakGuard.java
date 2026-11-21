package io.github.dornol.excelkit.testsupport;

import org.junit.jupiter.api.extension.AfterEachCallback;
import org.junit.jupiter.api.extension.BeforeEachCallback;
import org.junit.jupiter.api.extension.ExtensionContext;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.regex.Pattern;
import java.util.stream.Stream;

/**
 * Fails a test that leaves temp files of the library behind.
 * <p>
 * SXSSF keeps each sheet in a temp file under {@code java.io.tmpdir/poifiles} until its workbook is
 * closed. {@code ExcelHandler} closes the workbook in {@code writeTo}, and {@code ExcelWorkbook} /
 * {@code ExcelTemplateWriter} close it in {@code close()} unless {@code finish()} handed it to a
 * handler. A test that obtains a handler and never consumes it therefore leaks a file per sheet.
 * Consume it with {@code writeTo(OutputStream.nullOutputStream())}.
 * <p>
 * {@code CsvWriter} (and a few readers) stage their output in a directory named after a random UUID
 * directly under {@code java.io.tmpdir}; {@code CsvHandler.writeTo} removes it. Those directories
 * are counted too. JUnit's own {@code @TempDir} folders do not match the UUID pattern.
 * <p>
 * Registered through {@code META-INF/services} (auto-detection is enabled in
 * {@code junit-platform.properties}). The Gradle build gives every test JVM its own
 * {@code java.io.tmpdir}, so the count is not affected by other processes. Tests are not run in
 * parallel. Tests tagged {@code benchmark} and tests that already failed are not checked.
 */
public class TempFileLeakGuard implements BeforeEachCallback, AfterEachCallback {
    private static final ExtensionContext.Namespace NAMESPACE =
            ExtensionContext.Namespace.create(TempFileLeakGuard.class);
    private static final String BEFORE = "temp-files-before";
    private static final Pattern UUID_NAMED = Pattern.compile("[0-9a-f]{8}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{4}-[0-9a-f]{12}.*");

    @Override
    public void beforeEach(ExtensionContext context) throws IOException {
        context.getStore(NAMESPACE).put(BEFORE, leftoverCount());
    }

    @Override
    public void afterEach(ExtensionContext context) throws IOException {
        if (context.getTags().contains("benchmark") || context.getExecutionException().isPresent()) {
            return;
        }
        long before = context.getStore(NAMESPACE).get(BEFORE, Long.class);
        long leaked = leftoverCount() - before;
        if (leaked > 0) {
            throw new AssertionError(leaked + " temp file(s) or directorie(s) were left behind by this test. "
                    + "A workbook or CSV output was never released: consume the ExcelHandler / CsvHandler "
                    + "(for example writeTo(OutputStream.nullOutputStream())) or close the ExcelWorkbook / "
                    + "ExcelTemplateWriter.");
        }
    }

    private static long leftoverCount() throws IOException {
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
