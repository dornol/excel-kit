package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataValidation;
import org.apache.poi.ss.usermodel.DataValidationConstraint;
import org.apache.poi.ss.usermodel.DataValidationHelper;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.apache.poi.xssf.streaming.SXSSFSheet;
import org.jspecify.annotations.Nullable;

import java.time.LocalDate;
import java.util.Objects;

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
        LIST_FORMULA
    }

    /** Excel's limits for the title and the text of the error box. */
    private static final int MAX_ERROR_TITLE = 32;
    private static final int MAX_ERROR_MESSAGE = 255;

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

    private ExcelValidation(ValidationType type, int operator,
                            @Nullable String value1, @Nullable String value2) {
        this.type = type;
        this.operator = operator;
        this.value1 = value1;
        this.value2 = value2;
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
     * Applies this validation to the specified column range on the sheet.
     */
    void apply(DataValidationHelper helper, SXSSFSheet sheet, int col, int headerRow) {
        DataValidationConstraint constraint = createConstraint(helper);
        CellRangeAddressList range = new CellRangeAddressList(
                headerRow + 1, ExcelLimits.EXCEL_MAX_ROWS, col, col);
        DataValidation validation = helper.createValidation(constraint, range);
        validation.setShowErrorBox(showError);
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
        };
    }
}
