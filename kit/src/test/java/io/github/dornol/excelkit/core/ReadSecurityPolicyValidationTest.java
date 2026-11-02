package io.github.dornol.excelkit.core;

import org.junit.jupiter.api.Test;
import static org.junit.jupiter.api.Assertions.*;

class ReadSecurityPolicyValidationTest {
    @Test
    void bothConstructorsRejectNonFiniteAndNonPositiveCompressionRatios() {
        for (double ratio : new double[]{Double.NaN, Double.POSITIVE_INFINITY, Double.NEGATIVE_INFINITY, 0.0, -0.0, -1}) {
            assertThrows(IllegalArgumentException.class, () -> new ReadSecurityPolicy(true, true, 1024, 2048, ratio));
            assertThrows(IllegalArgumentException.class, () -> new ReadSecurityPolicy(true, true, 1024, 2048, ratio, 10));
        }
    }

    @Test
    void finitePositiveRatiosAndExistingPresetsRemainValid() {
        for (double ratio : new double[]{Double.MIN_VALUE, 0.5, 1.0, 100.0, Double.MAX_VALUE}) {
            assertEquals(ratio, new ReadSecurityPolicy(true, true, 1024, 2048, ratio).maxCompressionRatio());
            assertEquals(ratio, new ReadSecurityPolicy(true, true, 1024, 2048, ratio, 10).maxCompressionRatio());
        }
        assertNotNull(ReadSecurityPolicy.DEFAULT);
        assertNotNull(ReadSecurityPolicy.STRICT);
    }
}
