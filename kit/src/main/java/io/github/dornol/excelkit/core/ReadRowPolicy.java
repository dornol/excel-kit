package io.github.dornol.excelkit.core;

import java.util.List;

/** Applies common cancellation, blank-row, and emitted-row limits to a read stream. */
final class ReadRowPolicy {
    private final ReadOptions options;
    private long emittedRows;
    private long totalCells;
    private int consecutiveBlankRows;

    ReadRowPolicy(ReadOptions options) {
        this.options = options;
    }

    Decision evaluate(List<String> rawValues) {
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
        if (options.limits().maxTotalCells() >= 0) {
            totalCells += rawValues.size();
            if (totalCells > options.limits().maxTotalCells()) {
                throw new ReadLimitExceededException(ReadLimitExceededException.Limit.TOTAL_CELLS,
                        options.limits().maxTotalCells(), totalCells);
            }
        }
        if (options.maxRows() >= 0 && emittedRows >= options.maxRows()) {
            return Decision.STOP;
        }
        return Decision.PROCESS;
    }

    void recordEmitted() {
        emittedRows++;
    }

    enum Decision { PROCESS, SKIP, STOP }
}
