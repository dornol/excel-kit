package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.style.ExcelBorderStyle;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Constructor;
import java.lang.reflect.RecordComponent;
import java.util.HashSet;
import java.util.Set;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * Two cell styles that differ in any parameter must not share a cache key, or the second column would
 * silently get the first one's style. Walks every record component, so a new parameter that is left out
 * of {@link CellStyleParams#cacheKey()} fails here.
 */
class CellStyleParamsKeyTest {

    private static Object sample(Class<?> type, boolean other) {
        if (type == HorizontalAlignment.class) return other ? HorizontalAlignment.LEFT : HorizontalAlignment.RIGHT;
        if (type == VerticalAlignment.class) return other ? VerticalAlignment.TOP : VerticalAlignment.BOTTOM;
        if (type == ExcelBorderStyle.class) return other ? ExcelBorderStyle.THIN : ExcelBorderStyle.THICK;
        if (type == String.class) return other ? "a" : "b";
        if (type == Boolean.class) return other;
        if (type == Integer.class) return other ? 10 : 11;
        if (type == Short.class) return other ? (short) 1 : (short) 2;
        if (type == int[].class) return other ? new int[]{1, 2, 3} : new int[]{3, 2, 1};
        throw new AssertionError("add a sample for " + type);
    }

    private static CellStyleParams build(Constructor<?> constructor, RecordComponent[] components, int differing)
            throws Exception {
        Object[] args = new Object[components.length];
        for (int i = 0; i < args.length; i++) {
            args[i] = sample(components[i].getType(), i != differing);
        }
        return (CellStyleParams) constructor.newInstance(args);
    }

    @Test
    void everyParameterTakesPartInTheCacheKey() throws Exception {
        RecordComponent[] components = CellStyleParams.class.getRecordComponents();
        Class<?>[] types = new Class<?>[components.length];
        for (int i = 0; i < types.length; i++) types[i] = components[i].getType();
        Constructor<?> constructor = CellStyleParams.class.getDeclaredConstructor(types);
        constructor.setAccessible(true);

        Set<String> keys = new HashSet<>();
        keys.add(build(constructor, components, -1).cacheKey());
        for (int i = 0; i < components.length; i++) {
            String key = build(constructor, components, i).cacheKey();
            assertTrue(keys.add(key), "changing '" + components[i].getName() + "' does not change the cache key");
        }
        assertEquals(components.length + 1, keys.size());
    }
}
