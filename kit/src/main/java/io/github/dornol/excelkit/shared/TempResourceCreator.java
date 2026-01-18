package io.github.dornol.excelkit.shared;

import org.jspecify.annotations.NonNull;

import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.FileAttribute;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.util.Set;
import java.util.UUID;

/**
 * Utility class for creating temporary files and directories with appropriate permissions.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class TempResourceCreator {

    private TempResourceCreator() {
        /* empty */
    }

    /**
     * Creates a new temporary directory.
     * <p>
     * On POSIX-compliant file systems, the directory is created with restricted (rwx------) permissions.
     *
     * @return Path to the created temporary directory
     * @throws TempResourceCreateException If directory creation fails
     */
    @NonNull
    public static Path createTempDirectory() {
        try {
            if (FileSystems.getDefault().supportedFileAttributeViews().contains("posix")) {
                // 리눅스, 유닉스 계열
                Set<PosixFilePermission> perms = PosixFilePermissions.fromString("rwx------");
                FileAttribute<Set<PosixFilePermission>> attr = PosixFilePermissions.asFileAttribute(perms);
                return Files.createTempDirectory(UUID.randomUUID().toString(), attr);
            } else {
                // Windows
                return Files.createTempDirectory(UUID.randomUUID().toString());
            }
        } catch (IOException e) {
            throw new TempResourceCreateException(e);
        }
    }

    /**
     * Creates a new temporary file in the specified directory.
     *
     * @param directory The parent directory
     * @param prefix    The file prefix
     * @param suffix    The file suffix (extension)
     * @return Path to the created temporary file
     * @throws TempResourceCreateException If file creation fails
     */
    @NonNull
    public static Path createTempFile(Path directory, String prefix, String suffix) {
        try {
            return Files.createTempFile(directory, prefix, suffix);
        } catch (IOException e) {
            throw new TempResourceCreateException(e);
        }
    }

}
