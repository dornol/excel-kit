package io.github.dornol.excelkit.core;

@FunctionalInterface
public interface ReadProgressCallback {
    void onProgress(ReadProgress progress);
}
