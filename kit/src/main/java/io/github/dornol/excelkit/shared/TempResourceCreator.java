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

public class TempResourceCreator {

    private TempResourceCreator() {
        /* empty */
    }

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

    @NonNull
    public static Path createTempFile(Path directory, String prefix, String suffix) {
        try {
            return Files.createTempFile(directory, prefix, suffix);
        } catch (IOException e) {
            throw new TempResourceCreateException(e);
        }
    }

}
