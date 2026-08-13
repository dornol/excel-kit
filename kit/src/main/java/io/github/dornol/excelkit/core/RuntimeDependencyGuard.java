package io.github.dornol.excelkit.core;

/** Provides actionable errors when an optional format integration is missing at runtime. */
public final class RuntimeDependencyGuard {
    private RuntimeDependencyGuard() {
    }

    /**
     * Verifies that a runtime class is available and reports the required artifact when it is not.
     *
     * @param className class to resolve without initializing it
     * @param artifact Maven coordinates or dependency name to add
     * @param feature feature that requires the dependency
     * @throws ExcelKitException if the class cannot be loaded
     */
    public static void require(String className, String artifact, String feature) {
        try {
            Class.forName(className, false, RuntimeDependencyGuard.class.getClassLoader());
        } catch (ClassNotFoundException | LinkageError e) {
            throw new ExcelKitException(feature + " requires " + artifact
                    + " on the application's runtime classpath. Add " + artifact + " to your dependencies.", e);
        }
    }
}
