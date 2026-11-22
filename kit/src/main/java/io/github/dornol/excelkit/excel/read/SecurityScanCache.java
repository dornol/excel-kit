package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadSecurityException;
import io.github.dornol.excelkit.core.read.ReadSecurityPolicy;
import org.jspecify.annotations.Nullable;

import java.io.IOException;
import java.nio.file.Path;
import java.util.HashMap;
import java.util.Map;

/**
 * Remembers what the security scan of one workbook file found, per policy.
 * <p>
 * Reading several sheets of the same file with the same policy would otherwise scan the whole file (every
 * worksheet, when formulas are refused) once per sheet. The file and the policy fully determine the outcome, so
 * a clean result is reused, and so is a refusal: the next sheet is refused with the same exception without
 * scanning again. Only valid for one file that does not change while it is being read.
 */
final class SecurityScanCache {
    private final Map<ReadSecurityPolicy, @Nullable ReadSecurityException> outcomes = new HashMap<>();

    /**
     * Scans the file unless this policy already was.
     *
     * @throws ReadSecurityException if the policy refuses the workbook (now or when it was scanned before)
     * @throws IOException           if the file could not be scanned; that is not remembered, so the next sheet tries again
     */
    void scan(Path file, ReadSecurityPolicy policy) throws IOException {
        if (outcomes.containsKey(policy)) {
            ReadSecurityException refusal = outcomes.get(policy);
            if (refusal != null) throw refusal;
            return;
        }
        try {
            ExcelSecurityScanner.scan(file, policy);
            outcomes.put(policy, null);
        } catch (ReadSecurityException e) {
            outcomes.put(policy, e);
            throw e;
        }
    }
}
