package io.github.dornol.excelkit.core;

import java.util.List;

/** Applies common cancellation, blank-row, and emitted-row limits to a read stream. */
public final class ReadRowPolicy {
    private final ReadOptions options;
    private long emittedRows;
    private int consecutiveBlankRows;

    public ReadRowPolicy(ReadOptions options) {
        this.options = options;
    }

    public Decision evaluate(List<String> rawValues) {
        if (options.cancellationToken().isCancellationRequested()) {
            throw new ReadStoppedException();
        }
        if (rawValues.stream().allMatch(value -> value == null || value.isBlank())) {
            consecutiveBlankRows++;
            if (options.stopAtBlankRows() > 0
                    && consecutiveBlankRows >= options.stopAtBlankRows()) {
                return Decision.STOP;
            }
            if (options.skipBlankRows()) {
                return Decision.SKIP;
            }
        } else {
            consecutiveBlankRows = 0;
        }
        if (options.maxRows() >= 0 && emittedRows >= options.maxRows()) {
            return Decision.STOP;
        }
        return Decision.PROCESS;
    }

    public void recordEmitted() {
        emittedRows++;
    }

    public enum Decision { PROCESS, SKIP, STOP }
}
