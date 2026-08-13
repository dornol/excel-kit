package io.github.dornol.excelkit.core;

import org.junit.jupiter.api.Test;

import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;
import static org.junit.jupiter.api.Assertions.assertTrue;

class RuntimeDependencyGuardTest {
    @Test
    void reportsArtifactAndFeatureWhenDependencyIsMissing() {
        ExcelKitException error = assertThrows(ExcelKitException.class, () ->
                RuntimeDependencyGuard.require("missing.dependency.Type", "com.example:missing", "CSV reader"));

        assertTrue(error.getMessage().contains("CSV reader"));
        assertTrue(error.getMessage().contains("com.example:missing"));
        assertTrue(error.getMessage().contains("runtime classpath"));
    }

    @Test
    void acceptsAvailableDependency() {
        assertDoesNotThrow(() -> RuntimeDependencyGuard.require(
                "java.lang.String", "java.base", "core"));
    }
}
