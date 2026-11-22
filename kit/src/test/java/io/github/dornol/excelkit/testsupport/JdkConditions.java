package io.github.dornol.excelkit.testsupport;

/** Conditions on the running JDK for {@code @EnabledIf}. */
public final class JdkConditions {
    private JdkConditions() {
    }

    /**
     * Whether {@code java.util.zip.ZipInputStream} reads the ZIP64 entries POI writes in
     * {@code Zip64Mode.AlwaysWithCompatibility}: JDK 25 does, JDK 21 (and 17) reject them with "invalid entry size".
     */
    public static boolean zipInputStreamReadsZip64Compatibility() {
        return Runtime.version().feature() >= 25;
    }
}
