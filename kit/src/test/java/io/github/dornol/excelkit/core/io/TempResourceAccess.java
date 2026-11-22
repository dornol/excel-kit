package io.github.dornol.excelkit.core.io;

import org.jspecify.annotations.Nullable;

import java.nio.file.Path;

/**
 * Test-only bridge to {@link TempResourceContainer}'s protected members, for tests that live outside
 * the {@code core.io} package.
 */
public final class TempResourceAccess {

    private TempResourceAccess() {
    }

    public static @Nullable Path tempDir(TempResourceContainer container) {
        return container.getTempDir();
    }

    public static @Nullable Path tempFile(TempResourceContainer container) {
        return container.getTempFile();
    }

    public static void setTempDir(TempResourceContainer container, @Nullable Path dir) {
        container.setTempDir(dir);
    }

    public static void setTempFile(TempResourceContainer container, @Nullable Path file) {
        container.setTempFile(file);
    }
}
