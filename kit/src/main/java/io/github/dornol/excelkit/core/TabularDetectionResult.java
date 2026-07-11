package io.github.dornol.excelkit.core;

import org.jspecify.annotations.Nullable;
import java.nio.charset.Charset;

/** Detailed result of signature and text-sample inspection. */
public record TabularDetectionResult(TabularFileType type, DetectionConfidence confidence,
        @Nullable Charset charset, @Nullable Character delimiter) {}
