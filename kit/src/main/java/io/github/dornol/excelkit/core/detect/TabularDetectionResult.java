package io.github.dornol.excelkit.core.detect;

import org.jspecify.annotations.Nullable;
import java.nio.charset.Charset;

/**
 * Detailed result of signature and text-sample inspection.
 *
 * @since 0.21.0
 */
public record TabularDetectionResult(TabularFileType type, DetectionConfidence confidence,
        @Nullable Charset charset, @Nullable Character delimiter) {}
