package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.jspecify.annotations.Nullable;

import java.time.LocalDate;
import java.util.List;
import java.util.Map;
import java.util.Objects;
import java.util.function.ToIntFunction;

/**
 * Advanced data validation configuration for Excel columns.
 * <p>
 * Supports integer, decimal, text length, date, and custom formula validations.
 * Use the static factory methods to create validation instances.
 *
 * <pre>{@code
 * writer.column("Age", Person::getAge, c -> c
 *     .validation(ExcelValidation.integerBetween(1, 150)));
 * }</pre>
 *
 * @author dhkim
 * @since 0.7.0
 */
public class ExcelValidation {

    /**
     * Supported validation types.
     */
    public enum ValidationType {
        /** Integer value validation. */
        INTEGER,
        /** Decimal value validation. */
        DECIMAL,
        /** Text length validation. */
        TEXT_LENGTH,
        /** Date range validation. */
        DATE,
        /** Custom formula validation. */
        FORMULA,
        /** List from cell range validation. */
        LIST_FORMULA,
        /** List of fixed options. */
        LIST,
        /** List whose options follow the value of another column of the same row. */
        DEPENDENT_LIST
    }

    /** Excel's limits for the title and the text of the error box. */
    private static final int MAX_ERROR_TITLE = 32;
    private static final int MAX_ERROR_MESSAGE = 255;
    private static final int MAX_DROPDOWN_LENGTH = 255;

    private static void requireOrdered(double min, double max, String method) {
        if (min > max) {
            throw new IllegalArgumentException(method + " min " + min + " is above its max " + max
                    + ": no value could be valid");
        }
    }

    private static void requireText(@Nullable String text, String name) {
        Objects.requireNonNull(text, name + " cannot be null");
        if (text.isBlank()) throw new IllegalArgumentException(name + " cannot be blank");
    }

    private final ValidationType type;
    private final int operator;
    private final @Nullable String value1;
    private final @Nullable String value2;
    private @Nullable String errorTitle;
    private @Nullable String errorMessage;
    private boolean showError = true;
    private @Nullable String inputTitle;
    private @Nullable String inputMessage;
    private boolean allowBlank = true;

    private final String @Nullable [] listOptions;
    private final @Nullable DependentDropdown dependent;

    private ExcelValidation(ValidationType type, int operator,
                            @Nullable String value1, @Nullable String value2) {
        this(type, operator, value1, value2, null, null);
    }

    private ExcelValidation(ValidationType type, int operator, @Nullable String value1, @Nullable String value2,
                            String @Nullable [] listOptions, @Nullable DependentDropdown dependent) {
        this.type = type;
        this.operator = operator;
        this.value1 = value1;
        this.value2 = value2;
        this.listOptions = listOptions;
        this.dependent = dependent;
    }

    /**
     * Creates a validation that requires an integer between min and max (inclusive).
     *
     * @param min minimum value (inclusive)
     * @param max maximum value (inclusive)
     * @return the validation instance
     */
    public static ExcelValidation integerBetween(int min, int max) {
        requireOrdered(min, max, "integerBetween");
        return new ExcelValidation(ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.BETWEEN,
                String.valueOf(min), String.valueOf(max));
    }

    /**
     * Creates a validation that requires an integer greater than the given value.
     *
     * @param min the minimum threshold (exclusive)
     * @return the validation instance
     */
    public static ExcelValidation integerGreaterThan(int min) {
        return new ExcelValidation(ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.GREATER_THAN,
                String.valueOf(min), null);
    }

    /**
     * Creates a validation that requires an integer less than the given value.
     *
     * @param max the maximum threshold (exclusive)
     * @return the validation instance
     */
    public static ExcelValidation integerLessThan(int max) {
        return new ExcelValidation(ValidationType.INTEGER,
                DataValidationConstraint.OperatorType.LESS_THAN,
                String.valueOf(max), null);
    }

    /**
     * Creates a validation that requires a decimal between min and max (inclusive).
     *
     * @param min minimum value (inclusive)
     * @param max maximum value (inclusive)
     * @return the validation instance
     */
    public static ExcelValidation decimalBetween(double min, double max) {
        if (!Double.isFinite(min) || !Double.isFinite(max)) {
            throw new IllegalArgumentException("decimalBetween needs finite bounds, got " + min + " and " + max);
        }
        requireOrdered(min, max, "decimalBetween");
        return new ExcelValidation(ValidationType.DECIMAL,
                DataValidationConstraint.OperatorType.BETWEEN,
                String.valueOf(min), String.valueOf(max));
    }

    /**
     * Creates a validation that restricts text length between min and max (inclusive).
     *
     * @param min minimum text length (inclusive)
     * @param max maximum text length (inclusive)
     * @return the validation instance
     */
    public static ExcelValidation textLength(int min, int max) {
        if (min < 0) throw new IllegalArgumentException("textLength min must not be negative, got " + min);
        requireOrdered(min, max, "textLength");
        return new ExcelValidation(ValidationType.TEXT_LENGTH,
                DataValidationConstraint.OperatorType.BETWEEN,
                String.valueOf(min), String.valueOf(max));
    }

    /**
     * Creates a validation that restricts dates between start and end (inclusive).
     *
     * @param start the earliest allowed date
     * @param end the latest allowed date
     * @return the validation instance
     */
    public static ExcelValidation dateRange(LocalDate start, LocalDate end) {
        Objects.requireNonNull(start, "start cannot be null");
        Objects.requireNonNull(end, "end cannot be null");
        if (start.isAfter(end)) {
            throw new IllegalArgumentException("dateRange start " + start + " is after its end " + end);
        }
        return new ExcelValidation(ValidationType.DATE,
                DataValidationConstraint.OperatorType.BETWEEN,
                "DATE(" + start.getYear() + "," + start.getMonthValue() + "," + start.getDayOfMonth() + ")",
                "DATE(" + end.getYear() + "," + end.getMonthValue() + "," + end.getDayOfMonth() + ")");
    }

    /**
     * Creates a list validation that references a cell range for dropdown options.
     * <p>
     * Use this when dropdown options come from another sheet or cell range
     * instead of inline string arrays.
     *
     * @param range the cell range reference (e.g., "Sheet2!$A$1:$A$10")
     * @return the validation instance
     */
    public static ExcelValidation listFromRange(String range) {
        requireText(range, "range");
        return new ExcelValidation(ValidationType.LIST_FORMULA,
                DataValidationConstraint.OperatorType.BETWEEN, range, null);
    }

    /**
     * Creates a dropdown of fixed options, the same list as {@code dropdown(...)} on a column but with the options of
     * a validation: {@link #inputMessage(String) input message}, {@link #allowBlank(boolean) allowBlank} and the error
     * box.
     * <pre>{@code
     * .column("Status", p -> p.status(), c -> c.validation(
     *     ExcelValidation.list("Active", "Inactive").allowBlank(false).inputMessage("Pick one")))
     * }</pre>
     *
     * @param options the allowed values
     * @return the validation instance
     * @throws IllegalArgumentException if there is no option, an option contains a comma (use
     *         {@link #listFromRange(String)}), or the comma-joined options exceed Excel's 255-character limit
     * @since 0.29.0
     */
    public static ExcelValidation list(String... options) {
        return new ExcelValidation(ValidationType.LIST, DataValidationConstraint.OperatorType.BETWEEN, null, null,
                escapeListOptions(options), null);
    }

    /**
     * Creates a dropdown whose options depend on the value of another column of the same row, with the options of a
     * validation; see {@code ColumnStyleConfig#dependentDropdown(String, java.util.Map)} for how it works.
     *
     * @param parentColumn the name (header) of the parent column, another column of the same sheet
     * @param options      the options per parent value, in the order they are offered
     * @return the validation instance
     * @throws IllegalArgumentException if {@code parentColumn} is blank, {@code options} is empty, a parent value
     *         is empty, or a parent value has no option
     * @since 0.29.0
     */
    public static ExcelValidation dependentList(String parentColumn, Map<String, ? extends List<String>> options) {
        Objects.requireNonNull(options, "options cannot be null");
        return new ExcelValidation(ValidationType.DEPENDENT_LIST, DataValidationConstraint.OperatorType.BETWEEN,
                null, null, null, new DependentDropdown(parentColumn, new java.util.LinkedHashMap<>(options)));
    }

    /**
     * The options of an explicit list as Excel keeps them: a quote is written twice, and the joined text may not
     * pass 255 characters. Shared by {@code dropdown(...)} and {@link #list(String...)}.
     */
    static String[] escapeListOptions(String... options) {
        Objects.requireNonNull(options, "options cannot be null");
        if (options.length == 0) {
            throw new IllegalArgumentException("A dropdown needs at least one option");
        }
        String[] escaped = new String[options.length];
        for (int i = 0; i < options.length; i++) {
            String option = Objects.requireNonNull(options[i], "a dropdown option cannot be null");
            if (option.indexOf(',') >= 0) {
                throw new IllegalArgumentException("The dropdown option '" + option + "' contains a comma, which "
                        + "Excel reads as the end of an option; use ExcelValidation.listFromRange(...) for such values");
            }
            // inside the quoted list a quote is written twice
            escaped[i] = option.replace("\"", "\"\"");
        }
        int length = String.join(",", escaped).length();
        if (length > MAX_DROPDOWN_LENGTH) {
            throw new IllegalArgumentException("Dropdown options total " + length
                    + " characters (comma-joined); Excel allows at most " + MAX_DROPDOWN_LENGTH);
        }
        return escaped;
    }

    /**
     * Creates a validation using a custom Excel formula.
     *
     * @param formula the Excel formula
     * @return the validation instance
     */
    public static ExcelValidation formula(String formula) {
        requireText(formula, "formula");
        return new ExcelValidation(ValidationType.FORMULA,
                DataValidationConstraint.OperatorType.BETWEEN,
                formula, null);
    }

    /**
     * Sets the error dialog title.
     *
     * @param errorTitle the error dialog title
     * @return this instance for chaining
     */
    public ExcelValidation errorTitle(@Nullable String errorTitle) {
        if (errorTitle != null && errorTitle.length() > MAX_ERROR_TITLE) {
            throw new IllegalArgumentException("The error title is " + errorTitle.length()
                    + " characters; Excel allows at most " + MAX_ERROR_TITLE);
        }
        this.errorTitle = errorTitle;
        return this;
    }

    /**
     * Sets the error dialog message.
     *
     * @param errorMessage the error message text
     * @return this instance for chaining
     */
    public ExcelValidation errorMessage(@Nullable String errorMessage) {
        if (errorMessage != null && errorMessage.length() > MAX_ERROR_MESSAGE) {
            throw new IllegalArgumentException("The error message is " + errorMessage.length()
                    + " characters; Excel allows at most " + MAX_ERROR_MESSAGE);
        }
        this.errorMessage = errorMessage;
        return this;
    }

    /**
     * Sets whether to show the error dialog when validation fails.
     *
     * @param showError whether to show the error dialog
     * @return this instance for chaining
     */
    public ExcelValidation showError(boolean showError) {
        this.showError = showError;
        return this;
    }

    /**
     * Sets the title of the input message, the tooltip Excel shows when a cell of the column is selected.
     * The tooltip is shown when a title or a message is set.
     *
     * @param inputTitle the tooltip title (at most 32 characters), or {@code null} for none
     * @return this instance for chaining
     * @throws IllegalArgumentException if the title is over Excel's limit of 32 characters
     * @since 0.29.0
     */
    public ExcelValidation inputTitle(@Nullable String inputTitle) {
        if (inputTitle != null && inputTitle.length() > MAX_ERROR_TITLE) {
            throw new IllegalArgumentException("The input title is " + inputTitle.length()
                    + " characters; Excel allows at most " + MAX_ERROR_TITLE);
        }
        this.inputTitle = inputTitle;
        return this;
    }

    /**
     * Sets the input message, the tooltip Excel shows when a cell of the column is selected
     * (for example {@code "Enter the date as YYYY-MM-DD"}). The tooltip is shown when a title or a message is set.
     *
     * @param inputMessage the tooltip text (at most 255 characters), or {@code null} for none
     * @return this instance for chaining
     * @throws IllegalArgumentException if the message is over Excel's limit of 255 characters
     * @since 0.29.0
     */
    public ExcelValidation inputMessage(@Nullable String inputMessage) {
        if (inputMessage != null && inputMessage.length() > MAX_ERROR_MESSAGE) {
            throw new IllegalArgumentException("The input message is " + inputMessage.length()
                    + " characters; Excel allows at most " + MAX_ERROR_MESSAGE);
        }
        this.inputMessage = inputMessage;
        return this;
    }

    /**
     * Sets whether an empty cell passes the validation. It does by default; {@code allowBlank(false)} marks a
     * required column: Excel rejects clearing the cell by typing, but it does not flag a cell that was never filled,
     * so check required values when reading the upload.
     *
     * @param allowBlank {@code true} to let an empty cell pass
     * @return this instance for chaining
     * @since 0.29.0
     */
    public ExcelValidation allowBlank(boolean allowBlank) {
        this.allowBlank = allowBlank;
        return this;
    }

    /**
     * Applies this validation to the specified column range on the sheet.
     */
    void apply(DataValidationHelper helper, SXSSFSheet sheet, int col, int headerRow,
               ToIntFunction<String> parentColumnIndex) {
        DataValidationConstraint constraint = dependent != null
                ? helper.createFormulaListConstraint(dependent.formula(sheet,
                        parentColumnIndex.applyAsInt(dependent.parentColumn()), headerRow))
                : createConstraint(helper);
        CellRangeAddressList range = new CellRangeAddressList(
                headerRow + 1, ExcelLimits.EXCEL_MAX_ROWS, col, col);
        DataValidation validation = helper.createValidation(constraint, range);
        validation.setShowErrorBox(showError);
        validation.setEmptyCellAllowed(allowBlank);
        if (inputTitle != null || inputMessage != null) {
            validation.setShowPromptBox(true);
            validation.createPromptBox(inputTitle != null ? inputTitle : "", inputMessage != null ? inputMessage : "");
        }
        if (errorTitle != null) {
            validation.createErrorBox(
                    errorTitle,
                    errorMessage != null ? errorMessage : "");
        } else if (errorMessage != null) {
            validation.createErrorBox("Validation Error", errorMessage);
        }
        sheet.addValidationData(validation);
    }

    private DataValidationConstraint createConstraint(DataValidationHelper helper) {
        return switch (type) {
            case INTEGER -> helper.createIntegerConstraint(operator,
                    value1 != null ? value1 : "", value2);
            case DECIMAL -> helper.createDecimalConstraint(operator,
                    value1 != null ? value1 : "", value2);
            case TEXT_LENGTH -> helper.createTextLengthConstraint(operator,
                    value1 != null ? value1 : "", value2);
            case DATE -> helper.createDateConstraint(operator,
                    value1 != null ? value1 : "", value2,
                    "yyyy-MM-dd");
            case FORMULA -> helper.createCustomConstraint(value1 != null ? value1 : "TRUE");
            case LIST_FORMULA -> helper.createFormulaListConstraint(value1 != null ? value1 : "");
            case LIST -> helper.createExplicitListConstraint(listOptions != null ? listOptions : new String[0]);
            case DEPENDENT_LIST -> throw new IllegalStateException("a dependent list is built from its helper sheet");
        };
    }
}
