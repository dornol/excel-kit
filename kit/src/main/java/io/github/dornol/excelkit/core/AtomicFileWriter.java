package io.github.dornol.excelkit.core;

import java.io.IOException;
import java.io.OutputStream;
import java.nio.file.AtomicMoveNotSupportedException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.StandardCopyOption;
import java.nio.file.StandardOpenOption;
import java.util.UUID;

/**
 * Writes a file so that a failed write never destroys the previous content of the target.
 * <p>
 * The content goes to a temporary sibling file first and replaces the target only after the
 * write succeeded. If the write fails (or the handler was already consumed), the target is left
 * untouched and no partial file remains. Targets that are not regular files (for example
 * {@code /dev/stdout}) are written directly.
 *
 * @since 0.25.0
 */
public final class AtomicFileWriter {
    private AtomicFileWriter() {
    }

    /** Writes into the target through {@code writer}; see the class description for the guarantees. */
    public static void write(Path target, IoWriter writer) throws IOException {
        java.util.Objects.requireNonNull(target, "path cannot be null");
        Path destination = Files.isSymbolicLink(target) ? target.toRealPath() : target;
        if (Files.exists(destination) && !Files.isRegularFile(destination)) {
            try (OutputStream out = Files.newOutputStream(destination)) {
                writer.write(out);
            }
            return;
        }
        Path directory = destination.toAbsolutePath().getParent();
        Path temp = directory.resolve("." + destination.getFileName() + "." + UUID.randomUUID() + ".tmp");
        try {
            // Created with default permissions, like a directly written file.
            try (OutputStream out = Files.newOutputStream(temp, StandardOpenOption.CREATE_NEW,
                    StandardOpenOption.WRITE)) {
                writer.write(out);
            }
            try {
                Files.move(temp, destination, StandardCopyOption.REPLACE_EXISTING, StandardCopyOption.ATOMIC_MOVE);
            } catch (AtomicMoveNotSupportedException e) {
                Files.move(temp, destination, StandardCopyOption.REPLACE_EXISTING);
            }
        } finally {
            Files.deleteIfExists(temp);
        }
    }

    /** Produces the file content. */
    @FunctionalInterface
    public interface IoWriter {
        /** Writes the content to {@code out}; the stream is closed by the caller. */
        void write(OutputStream out) throws IOException;
    }
}
