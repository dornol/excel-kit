package io.github.dornol.excelkit.shared;

import org.jspecify.annotations.NonNull;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import java.io.IOException;
import java.nio.file.FileSystems;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.attribute.AclEntry;
import java.nio.file.attribute.AclEntryPermission;
import java.nio.file.attribute.AclEntryType;
import java.nio.file.attribute.AclFileAttributeView;
import java.nio.file.attribute.FileAttribute;
import java.nio.file.attribute.PosixFilePermission;
import java.nio.file.attribute.PosixFilePermissions;
import java.nio.file.attribute.UserPrincipal;
import java.util.EnumSet;
import java.util.List;
import java.util.Set;
import java.util.UUID;

/**
 * Utility class for creating temporary files and directories with appropriate permissions.
 *
 * @author dhkim
 * @since 2025-07-19
 */
public class TempResourceCreator {
    private static final Logger log = LoggerFactory.getLogger(TempResourceCreator.class);

    private TempResourceCreator() {
        /* empty */
    }

    /**
     * Creates a new temporary directory.
     * <p>
     * On POSIX-compliant file systems, the directory is created with restricted (rwx------) permissions.
     * On Windows (ACL-based), the directory's ACL is restricted to the current user only.
     *
     * @return Path to the created temporary directory
     * @throws TempResourceCreateException If directory creation fails
     */
    @NonNull
    public static Path createTempDirectory() {
        try {
            if (FileSystems.getDefault().supportedFileAttributeViews().contains("posix")) {
                Set<PosixFilePermission> perms = PosixFilePermissions.fromString("rwx------");
                FileAttribute<Set<PosixFilePermission>> attr = PosixFilePermissions.asFileAttribute(perms);
                return Files.createTempDirectory(UUID.randomUUID().toString(), attr);
            } else {
                Path dir = Files.createTempDirectory(UUID.randomUUID().toString());
                restrictToOwnerOnWindows(dir);
                return dir;
            }
        } catch (IOException e) {
            throw new TempResourceCreateException(e);
        }
    }

    /**
     * Restricts the given path's ACL to the current user only (Windows).
     * If ACL modification fails, a warning is logged but no exception is thrown.
     */
    private static void restrictToOwnerOnWindows(Path path) {
        try {
            AclFileAttributeView aclView = Files.getFileAttributeView(path, AclFileAttributeView.class);
            if (aclView == null) {
                return;
            }
            UserPrincipal owner = aclView.getOwner();
            AclEntry entry = AclEntry.newBuilder()
                    .setType(AclEntryType.ALLOW)
                    .setPrincipal(owner)
                    .setPermissions(EnumSet.allOf(AclEntryPermission.class))
                    .build();
            aclView.setAcl(List.of(entry));
        } catch (IOException e) {
            log.warn("Failed to restrict temp directory ACL: {}", path, e);
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
