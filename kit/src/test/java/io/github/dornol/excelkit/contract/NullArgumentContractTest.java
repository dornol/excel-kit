package io.github.dornol.excelkit.contract;

import io.github.dornol.excelkit.csv.CsvHandler;
import io.github.dornol.excelkit.csv.CsvReader;
import io.github.dornol.excelkit.csv.CsvWriter;
import io.github.dornol.excelkit.excel.ColumnConfig;
import io.github.dornol.excelkit.excel.ExcelChartConfig;
import io.github.dornol.excelkit.excel.ExcelConditionalRule;
import io.github.dornol.excelkit.excel.ExcelHandler;
import io.github.dornol.excelkit.excel.ExcelPrintSetup;
import io.github.dornol.excelkit.excel.ExcelReader;
import io.github.dornol.excelkit.excel.ExcelSummary;
import io.github.dornol.excelkit.excel.ExcelWorkbook;
import io.github.dornol.excelkit.excel.ExcelWriter;
import io.github.dornol.excelkit.excel.HeaderStyleConfig;
import io.github.dornol.excelkit.excel.RowStyleConfig;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.OutputStream;
import java.lang.annotation.Annotation;
import java.lang.reflect.Array;
import java.lang.reflect.Constructor;
import java.lang.reflect.InvocationTargetException;
import java.lang.reflect.Method;
import java.lang.reflect.Modifier;
import java.lang.reflect.Parameter;
import java.lang.reflect.Proxy;
import java.nio.charset.StandardCharsets;
import java.nio.file.Path;
import java.time.Duration;
import java.util.ArrayList;
import java.util.Collection;
import java.util.LinkedHashMap;
import java.util.LinkedHashSet;
import java.util.List;
import java.util.Locale;
import java.util.Map;
import java.util.Set;
import java.util.TreeSet;
import java.util.function.Supplier;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.assertTrue;

/**
 * The null contract of the public API: a parameter that is not annotated {@code @Nullable} rejects
 * {@code null} where it is passed, with a {@link NullPointerException} or {@link IllegalArgumentException},
 * not later when a row is written or read (or never).
 * <p>
 * The test walks the public methods of the fluent entry points by reflection and calls each one with
 * {@code null} for every non-{@code @Nullable} reference parameter in turn, all other parameters being
 * valid sample values. A new setter is therefore checked without anyone remembering to add a test.
 * A parameter that really accepts {@code null} without saying so is listed in {@link #ACCEPTS_NULL}
 * with the reason; the right fix is usually to annotate it {@code @Nullable} instead.
 */
class NullArgumentContractTest {

    /** A receiver of the calls, and what to release afterwards. */
    private record Receiver(Object target, AutoCloseable cleanup) {
    }

    private record Target(Class<?> type, Supplier<Receiver> receiver) {
    }

    /** Parameters that accept {@code null} on purpose although they are not annotated: key to reason. */
    private static final Map<String, String> ACCEPTS_NULL = Map.of();

    private static final AutoCloseable NOTHING = () -> { };

    private static final List<Target> TARGETS = List.of(
            new Target(ExcelWriter.class, () -> new Receiver(ExcelWriter.<String>create(), NOTHING)),
            new Target(ExcelWorkbook.class, () -> {
                ExcelWorkbook workbook = ExcelWorkbook.create();
                return new Receiver(workbook, workbook);
            }),
            new Target(io.github.dornol.excelkit.excel.ExcelSheetWriter.class, () -> {
                ExcelWorkbook workbook = ExcelWorkbook.create();
                return new Receiver(workbook.<String>sheet("S"), workbook);
            }),
            new Target(CsvWriter.class, () -> new Receiver(CsvWriter.<String>create(), NOTHING)),
            new Target(CsvReader.class, () -> new Receiver(CsvReader.setter(Object::new), NOTHING)),
            new Target(ExcelReader.class, () -> new Receiver(ExcelReader.setter(Object::new), NOTHING)),
            new Target(ColumnConfig.class, () -> new Receiver(construct(ColumnConfig.class), NOTHING)),
            new Target(RowStyleConfig.class, () -> new Receiver(construct(RowStyleConfig.class), NOTHING)),
            new Target(HeaderStyleConfig.class, () -> new Receiver(construct(HeaderStyleConfig.class), NOTHING)),
            new Target(ExcelSummary.class, () -> new Receiver(construct(ExcelSummary.class), NOTHING)),
            new Target(ExcelConditionalRule.class, () -> new Receiver(construct(ExcelConditionalRule.class), NOTHING)),
            new Target(ExcelChartConfig.class, () -> new Receiver(construct(ExcelChartConfig.class), NOTHING)),
            new Target(ExcelPrintSetup.class, () -> new Receiver(construct(ExcelPrintSetup.class), NOTHING)));

    /** Methods that are not part of the contract: they take the object itself or are Object's. */
    private static boolean isContractMethod(Method method) {
        if (method.getDeclaringClass() == Object.class) return false;
        if (method.isSynthetic() && !method.isBridge()) return false;
        return Modifier.isPublic(method.getModifiers());
    }

    @Test
    void everyNonNullableParameterRejectsNull() {
        List<String> violations = new ArrayList<>();
        Set<String> unsupported = new TreeSet<>();
        int checked = 0;

        for (Target target : TARGETS) {
            for (Method method : publicMethods(target.type())) {
                Parameter[] parameters = method.getParameters();
                for (int position = 0; position < parameters.length; position++) {
                    if (!mustRejectNull(method, parameters[position], position)) continue;
                    String key = describe(target.type(), method, position);
                    if (ACCEPTS_NULL.containsKey(key)) continue;

                    Object[] arguments = sampleArguments(method);
                    if (arguments == null) {
                        unsupported.add(key);
                        continue;
                    }
                    arguments[position] = null;
                    checked++;
                    String outcome = call(target, method, arguments);
                    if (outcome != null) violations.add(key + " -> " + outcome);
                }
            }
        }

        System.out.println("[null contract] checked " + checked + " parameters, "
                + unsupported.size() + " not checkable (no sample value): " + unsupported);
        assertTrue(checked > 150, "the walk found too few parameters (" + checked + "); is it still reflecting?");
        assertTrue(violations.isEmpty(), violations.size() + " parameter(s) accept null without saying so:\n  "
                + String.join("\n  ", violations));
    }

    // ---- what to check

    private static boolean mustRejectNull(Method method, Parameter parameter, int position) {
        if (parameter.getType().isPrimitive()) return false;
        if (method.isVarArgs() && position == method.getParameterCount() - 1) return false;
        return !isNullable(parameter);
    }

    private static boolean isNullable(Parameter parameter) {
        for (Annotation annotation : parameter.getAnnotations()) {
            if (annotation.annotationType().getSimpleName().equals("Nullable")) return true;
        }
        for (Annotation annotation : parameter.getAnnotatedType().getAnnotations()) {
            if (annotation.annotationType().getSimpleName().equals("Nullable")) return true;
        }
        return false;
    }

    private static List<Method> publicMethods(Class<?> type) {
        Map<String, Method> bySignature = new LinkedHashMap<>();
        for (Method method : type.getMethods()) {
            if (!isContractMethod(method)) continue;
            String signature = method.getName() + java.util.Arrays.toString(method.getParameterTypes());
            Method known = bySignature.get(signature);
            if (known == null || (known.isBridge() && !method.isBridge())) bySignature.put(signature, method);
        }
        return new ArrayList<>(bySignature.values());
    }

    private static String describe(Class<?> type, Method method, int position) {
        StringBuilder text = new StringBuilder(type.getSimpleName()).append('#').append(method.getName()).append('(');
        Class<?>[] types = method.getParameterTypes();
        for (int i = 0; i < types.length; i++) {
            if (i > 0) text.append(", ");
            text.append(types[i].getSimpleName());
            if (i == position) text.append(" <null>");
        }
        return text.append(')').toString();
    }

    // ---- the call

    /** {@code null} when the call rejected the null as it should, otherwise what happened instead. */
    private static String call(Target target, Method method, Object[] arguments) {
        Receiver receiver = Modifier.isStatic(method.getModifiers()) ? null : target.receiver().get();
        try {
            method.setAccessible(true);
            Object result = method.invoke(receiver == null ? null : receiver.target(), arguments);
            release(result);
            return "accepted the null (returned " + (result == null ? "null" : result.getClass().getSimpleName()) + ")";
        } catch (InvocationTargetException e) {
            Throwable cause = e.getCause();
            if (cause instanceof NullPointerException || cause instanceof IllegalArgumentException) return null;
            return "failed with " + cause.getClass().getSimpleName() + ": " + cause.getMessage();
        } catch (ReflectiveOperationException | RuntimeException e) {
            return "could not be called: " + e;
        } finally {
            if (receiver != null) {
                try {
                    receiver.cleanup().close();
                } catch (Exception ignored) {
                    // best effort
                }
            }
        }
    }

    private static void release(Object result) {
        try {
            if (result instanceof ExcelHandler handler) handler.discard();
            else if (result instanceof CsvHandler handler) handler.discard();
            else if (result instanceof AutoCloseable closeable) closeable.close();
        } catch (Exception ignored) {
            // best effort
        }
    }

    // ---- sample values

    private static Object[] sampleArguments(Method method) {
        Parameter[] parameters = method.getParameters();
        Object[] arguments = new Object[parameters.length];
        for (int i = 0; i < parameters.length; i++) {
            Object sample = sample(parameters[i].getType());
            if (sample == UNSUPPORTED) return null;
            arguments[i] = sample;
        }
        return arguments;
    }

    private static final Object UNSUPPORTED = new Object();

    private static Object sample(Class<?> type) {
        if (type == int.class || type == Integer.class) return 1;
        if (type == long.class || type == Long.class) return 1L;
        if (type == short.class) return (short) 1;
        if (type == byte.class) return (byte) 1;
        if (type == float.class || type == Float.class) return 1f;
        if (type == double.class || type == Double.class) return 1d;
        if (type == boolean.class || type == Boolean.class) return true;
        if (type == char.class) return ',';
        if (type == String.class || type == CharSequence.class || type == Object.class) return "x";
        if (type == char[].class) return new char[]{'x'};
        if (type.isArray()) {
            Object element = sample(type.getComponentType());
            if (element == UNSUPPORTED) return UNSUPPORTED;
            Object array = Array.newInstance(type.getComponentType(), 1);
            Array.set(array, 0, element);
            return array;
        }
        if (type.isEnum()) return type.getEnumConstants()[0];
        if (type == java.io.InputStream.class) return new ByteArrayInputStream(new byte[0]);
        if (type == OutputStream.class) return OutputStream.nullOutputStream();
        if (type == Path.class) return Path.of("contract-sample");
        if (type == java.nio.charset.Charset.class) return StandardCharsets.UTF_8;
        if (type == Locale.class) return Locale.ROOT;
        if (type == Duration.class) return Duration.ofSeconds(1);
        if (type == Stream.class) return Stream.empty();
        if (type == List.class || type == Collection.class || type == Iterable.class) return List.of("x");
        if (type == Set.class) return Set.of("x");
        if (type == Map.class) return Map.of("x", "x");
        Object constant = staticConstant(type);
        if (constant != null) return constant;
        Object produced = staticFactory(type);
        if (produced != null) return produced;
        if (type.isInterface()) return proxy(type);
        Object built = construct(type);
        return built != null ? built : UNSUPPORTED;
    }

    /** A public static final field of the type itself, such as {@code ExcelColor.RED}. */
    private static Object staticConstant(Class<?> type) {
        for (java.lang.reflect.Field field : type.getFields()) {
            if (Modifier.isStatic(field.getModifiers()) && Modifier.isFinal(field.getModifiers())
                    && field.getType() == type) {
                try {
                    Object value = field.get(null);
                    if (value != null) return value;
                } catch (IllegalAccessException ignored) {
                    // try the next one
                }
            }
        }
        return null;
    }

    /**
     * A public static method of the type that returns the type from no argument or one String
     * ({@code CellConversionConfig.defaults()}, {@code TableOptions.defaults("x")}), or a static
     * {@code builder()} whose {@code build()} returns the type ({@code ReaderOptions}).
     */
    private static Object staticFactory(Class<?> type) {
        for (Method method : type.getMethods()) {
            if (!Modifier.isStatic(method.getModifiers())) continue;
            Class<?>[] parameters = method.getParameterTypes();
            try {
                if (method.getReturnType() == type && (parameters.length == 0
                        || (parameters.length == 1 && parameters[0] == String.class))) {
                    Object value = method.invoke(null, parameters.length == 0 ? new Object[0] : new Object[]{"x"});
                    if (value != null) return value;
                } else if (method.getName().equals("builder") && parameters.length == 0) {
                    Object builder = method.invoke(null);
                    Method build = builder.getClass().getMethod("build");
                    if (build.getReturnType() == type) return build.invoke(builder);
                }
            } catch (ReflectiveOperationException | RuntimeException ignored) {
                // try the next one
            }
        }
        return null;
    }

    private static Object proxy(Class<?> type) {
        return Proxy.newProxyInstance(type.getClassLoader(), new Class<?>[]{type}, (proxy, method, args) -> {
            Class<?> returned = method.getReturnType();
            if (returned == boolean.class) return false;
            if (returned == int.class) return 0;
            if (returned == long.class) return 0L;
            if (returned == double.class) return 0d;
            if (returned == float.class) return 0f;
            if (returned == char.class) return '\0';
            if (returned == short.class) return (short) 0;
            if (returned == byte.class) return (byte) 0;
            return null;
        });
    }

    private static <T> T construct(Class<T> type) {
        try {
            Constructor<T> constructor = type.getDeclaredConstructor();
            constructor.setAccessible(true);
            return constructor.newInstance();
        } catch (ReflectiveOperationException | RuntimeException e) {
            return null;
        }
    }
}
