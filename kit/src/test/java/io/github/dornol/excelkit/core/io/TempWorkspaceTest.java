package io.github.dornol.excelkit.core.io;

import org.junit.jupiter.api.Test;

import java.nio.file.Files;
import java.nio.file.Path;

import static org.junit.jupiter.api.Assertions.*;

class TempWorkspaceTest {

    @Test
    void closingRemovesTheDirectoryAndEveryFileInIt() throws Exception {
        Path directory;
        Path first;
        Path second;
        try (TempWorkspace workspace = TempWorkspace.create()) {
            directory = workspace.directory();
            first = workspace.newFile("a-", ".xlsx");
            second = workspace.newFile("b-", ".tmp");
            Files.writeString(first, "x");
            Files.createDirectory(directory.resolve("nested"));
            Files.writeString(directory.resolve("nested").resolve("c"), "y");
            assertTrue(Files.exists(second));
            assertEquals(directory, first.getParent());
        }

        assertFalse(Files.exists(directory));
    }

    @Test
    void closingTwiceIsHarmless() {
        TempWorkspace workspace = TempWorkspace.create();
        workspace.close();
        assertDoesNotThrow(workspace::close);
    }

    @Test
    void theDirectoryIsPrivateToTheOwnerWhereTheFileSystemHasPermissions() throws Exception {
        try (TempWorkspace workspace = TempWorkspace.create()) {
            if (!workspace.directory().getFileSystem().supportedFileAttributeViews().contains("posix")) return;
            assertEquals("rwx------", java.nio.file.attribute.PosixFilePermissions.toString(
                    Files.getPosixFilePermissions(workspace.directory())));
        }
    }
}
