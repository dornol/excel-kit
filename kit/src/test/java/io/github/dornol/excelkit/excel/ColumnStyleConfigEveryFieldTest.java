package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelCellComment;
import io.github.dornol.excelkit.excel.style.CellColorFunction;
import org.junit.jupiter.api.Test;

import java.lang.reflect.Field;
import java.lang.reflect.Modifier;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Objects;
import java.util.function.Function;

import static org.junit.jupiter.api.Assertions.*;

/**
 * A setting of a column is a field of {@link ColumnStyleConfig}, and a writer-level default has to reach a column that
 * did not set it ({@code applyDefaults}). This goes over the fields themselves, so a field added later without its line
 * in {@code applyDefaults} fails here instead of silently ignoring the default.
 */
class ColumnStyleConfigEveryFieldTest {

    /** A value of the field's type that is not the field's initial one. */
    private static Object valueFor(Field field) {
        Class<?> type = field.getType();
        String name = field.getName();
        if (name.equals("validation")) return ExcelValidation.list("a", "b");
        if (name.equals("headerComment")) return ExcelCellComment.of("note");
        if (type == boolean.class || type == Boolean.class) return true;
        if (type == int.class || type == Integer.class) return 7;
        if (type == Short.class || type == short.class) return (short) 7;
        if (type == String.class) return "x";
        if (type == Object.class) return "nothing";
        if (type == int[].class) return new int[]{1, 2, 3};
        if (type == String[].class) return new String[]{"a", "b"};
        if (type == Function.class) return (Function<Object, String>) o -> "c";
        if (type == CellColorFunction.class) return (CellColorFunction<Object>) (v, row) -> null;
        if (type.isEnum()) {
            Object[] constants = type.getEnumConstants();
            return constants[constants.length - 1];
        }
        throw new AssertionError("no test value for " + field + ": add one here");
    }

    private static boolean sameValue(Object expected, Object actual) {
        if (expected instanceof int[] a) return Arrays.equals(a, (int[]) actual);
        if (expected instanceof Object[] a) return Arrays.equals(a, (Object[]) actual);
        return Objects.equals(expected, actual);
    }

    private static List<Field> settings() {
        List<Field> fields = new ArrayList<>();
        for (Field field : ColumnStyleConfig.class.getDeclaredFields()) {
            if (Modifier.isStatic(field.getModifiers()) || field.isSynthetic()) continue;
            fields.add(field);
        }
        return fields;
    }

    @Test
    void everySettingOfADefaultReachesAColumnThatDidNotSetIt() throws Exception {
        List<Field> fields = settings();
        assertTrue(fields.size() > 30, "the fields were found: " + fields.size());
        for (Field field : fields) {
            field.setAccessible(true);
            if (field.getName().equals("alignment")) continue; // follows alignmentSet, below
            var defaults = new ColumnStyleConfig.DefaultStyleConfig<Object>();
            Object value = valueFor(field);
            field.set(defaults, value);
            if (field.getName().equals("alignmentSet")) {
                defaults.alignment(org.apache.poi.ss.usermodel.HorizontalAlignment.RIGHT);
                value = true;
            }
            var column = new ColumnConfig<Object>();
            column.applyDefaults(defaults);
            assertTrue(sameValue(value, field.get(column)),
                    "the default of '" + field.getName() + "' did not reach the column: add it to applyDefaults");
        }
    }

    @Test
    void aSettingOfTheColumnIsNotOverriddenByTheDefault() throws Exception {
        for (Field field : settings()) {
            field.setAccessible(true);
            String name = field.getName();
            Class<?> type = field.getType();
            // Flags that only switch on (a default may switch them on, the column cannot switch them off again),
            // the alignment pair, and the dropdown/validation pair (a cell has one validation) are not "column wins".
            if (type == boolean.class || name.startsWith("alignment") || name.equals("dropdownOptions")
                    || name.equals("validation") || name.equals("commentWidth") || name.equals("commentHeight")
                    || name.equals("minWidth") || name.equals("maxWidth")) continue;
            var defaults = new ColumnStyleConfig.DefaultStyleConfig<Object>();
            var column = new ColumnConfig<Object>();
            Object own = valueFor(field);
            field.set(column, own);
            Field defaultsField = field;
            Object other = valueFor(field);
            if (other instanceof Boolean b) other = !b;
            else if (other instanceof Integer i) other = i + 1;
            else if (other instanceof Short s) other = (short) (s + 1);
            else if (other instanceof String) other = "other";
            else if (other instanceof Object[]) other = new String[]{"z"};
            else if (other instanceof int[]) other = new int[]{9, 9, 9};
            else if (type.isEnum()) {
                Object[] constants = type.getEnumConstants();
                if (constants.length > 1) other = constants[0];
            }
            defaultsField.set(defaults, other);
            column.applyDefaults(defaults);
            assertTrue(sameValue(own, field.get(column)), "'" + name + "' of the column was overridden by the default");
        }
    }
}
