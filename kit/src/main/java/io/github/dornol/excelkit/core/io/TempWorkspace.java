package io.github.dornol.excelkit.core.io;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.util.Comparator;
import java.util.stream.Stream;

/**
 * A private temporary directory that is removed, with everything in it, when the workspace is closed.
 * <p>
 * The directory is created with owner-only permissions (see {@link TempResourceCreator#createTempDirectory()}).
 * Closing deletes the files that were created in it and then the directory; a path that cannot be deleted is logged
 * and left to be deleted when the JVM exits, so cleanup never hides the exception that is being thrown.
 *
 * <pre>{@code
 * try (TempWorkspace workspace = TempWorkspace.create()) {
 *     Path copy = workspace.newFile("upload-", ".xlsx");
 *     // ... use copy ...
 * }   // copy and its directory are gone
 * }</pre>
 *
 * @since 0.28.0
 */
public final class TempWorkspace implements AutoCloseable {
    private static final Logger log = LoggerFactory.getLogger(TempWorkspace.class);

    private final Path directory;

    private TempWorkspace(Path directory) {
        this.directory = directory;
    }

    /**
     * Creates the directory.
     *
     * @return a new workspace
     * @throws TempResourceCreateException if the directory cannot be created
     */
    public static TempWorkspace create() {
        return new TempWorkspace(TempResourceCreator.createTempDirectory());
    }

    /**
     * The directory this workspace owns.
     *
     * @return the directory
     */
    public Path directory() {
        return directory;
    }

    /**
     * Creates an empty file in the workspace.
     *
     * @param prefix the start of the file name
     * @param suffix the end of the file name, such as {@code ".xlsx"}
     * @return the new file
     * @throws TempResourceCreateException if the file cannot be created
     */
    public Path newFile(String prefix, String suffix) {
        return TempResourceCreator.createTempFile(directory, prefix, suffix);
    }

    /** Deletes the directory and everything in it. Closing twice is harmless. */
    @Override
    public void close() {
        if (!Files.exists(directory)) return;
        try (Stream<Path> paths = Files.walk(directory)) {
            // Deepest first, so a directory is empty by the time it is deleted.
            paths.sorted(Comparator.reverseOrder()).forEach(TempWorkspace::deleteQuietly);
        } catch (IOException e) {
            log.warn("Failed to clean up temp directory: {}", directory, e);
            directory.toFile().deleteOnExit();
        }
    }

    /** Deletes one path if it exists; one that stays is logged and left to the JVM's exit. */
    static void deleteQuietly(Path path) {
        try {
            Files.deleteIfExists(path);
        } catch (IOException e) {
            log.warn("Failed to delete temp path: {}", path, e);
            path.toFile().deleteOnExit();
        }
    }
}
