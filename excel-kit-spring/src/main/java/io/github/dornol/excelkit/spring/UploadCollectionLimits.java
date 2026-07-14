package io.github.dornol.excelkit.spring;

/** Bounds in-memory upload result collections while preserving full summary counts. */
public record UploadCollectionLimits(int maxSuccessRows, int maxErrors) {
    public static final UploadCollectionLimits UNLIMITED = new UploadCollectionLimits(-1, -1);
    public UploadCollectionLimits {
        if (maxSuccessRows < -1 || maxErrors < -1) throw new IllegalArgumentException("limits must be >= -1");
    }
}
