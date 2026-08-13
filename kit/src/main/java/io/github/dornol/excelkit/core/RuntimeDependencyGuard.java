package io.github.dornol.excelkit.core;

/** Provides actionable errors when an optional format integration is missing at runtime. */
public final class RuntimeDependencyGuard {
    private RuntimeDependencyGuard() {
    }

    public static void require(String className, String artifact, String feature) {
        try {
            Class.forName(className, false, RuntimeDependencyGuard.class.getClassLoader());
        } catch (ClassNotFoundException | LinkageError e) {
            throw new ExcelKitException(feature + " requires " + artifact
                    + " on the application's runtime classpath. Add " + artifact + " to your dependencies.", e);
        }
    }
}
