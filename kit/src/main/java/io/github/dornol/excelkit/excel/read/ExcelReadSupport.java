package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.CellConversionConfig;
import io.github.dornol.excelkit.core.io.TempWorkspace;
import org.apache.poi.poifs.crypt.Decryptor;
import org.apache.poi.poifs.crypt.EncryptionInfo;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;
import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.util.LocaleUtil;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.apache.poi.xssf.eventusermodel.XSSFSheetXMLHandler;
import org.apache.poi.xssf.model.SharedStrings;
import org.apache.poi.xssf.model.Styles;
import org.apache.poi.xssf.usermodel.XSSFRichTextString;
import org.xml.sax.InputSource;
import org.xml.sax.Attributes;
import org.xml.sax.SAXException;
import org.xml.sax.helpers.AttributesImpl;
import org.xml.sax.helpers.DefaultHandler;
import org.jspecify.annotations.Nullable;

import java.io.InputStream;
import java.io.OutputStream;
import java.nio.file.Files;
import java.nio.file.Path;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
import java.util.Calendar;
import java.util.Locale;
import java.util.function.Supplier;
import java.util.regex.Pattern;

/**
 * Package-private utility methods shared by {@link ExcelReadHandler} and {@link ExcelReader}
 * to eliminate duplicate read logic.
 *
 * @author dhkim
 */
class ExcelReadSupport {

    /**
     * Decrypts an agile-encrypted workbook into {@code target}.
     *
     * @throws ExcelReadException if the password is wrong
     */
    static Path decryptInto(TempWorkspace workspace, Path encrypted, String password) throws Exception {
        Path target = workspace.newFile("decrypted-", ".xlsx");
        decrypt(encrypted, password, target);
        return target;
    }

    static void decrypt(Path encrypted, String password, Path target) throws Exception {
        try (POIFSFileSystem fs =
                     new POIFSFileSystem(encrypted.toFile(), true)) {
            EncryptionInfo info = new EncryptionInfo(fs);
            Decryptor decryptor = Decryptor.getInstance(info);
            if (!decryptor.verifyPassword(password)) {
                throw new ExcelReadException("Invalid password for encrypted Excel file");
            }
            try (InputStream decrypted = decryptor.getDataStream(fs);
                 OutputStream out = Files.newOutputStream(target)) {
                decrypted.transferTo(out);
            }
        }
    }


    /** Excel maximum column count (XFD = 16,384). */
    static final int EXCEL_MAX_COLUMNS = 16_384;
    private static final Pattern FORMAT_LITERAL = Pattern.compile("\"[^\"]*\"|\\\\.|_.|\\*.");
    private static final Pattern ELAPSED_TIME = Pattern.compile("\\[(?:h+|m+|s+)\\]");
    private static final Pattern FORMAT_MODIFIER = Pattern.compile("\\[[^\\]]*\\]");

    private ExcelReadSupport() {
    }

    /**
     * Uses the workbook's date system for numeric cells and cached formula results, and the default conversion locale
     * (the root locale), not the machine's.
     */
    static DataFormatter dataFormatter(XSSFReader reader) throws Exception {
        return dataFormatter(reader, CellConversionConfig.defaults().locale());
    }

    static DataFormatter dataFormatter(XSSFReader reader, Locale locale) throws Exception {
        boolean[] date1904 = {false};
        var parser = XMLHelper.newXMLReader();
        parser.setContentHandler(new DefaultHandler() {
            @Override
            public void startElement(String uri, String localName, String qName, Attributes attributes) {
                if ("workbookPr".equals(localName) || "workbookPr".equals(qName)) {
                    String value = attributes.getValue("date1904");
                    date1904[0] = "1".equals(value) || "true".equals(value);
                }
            }
        });
        try (var workbook = reader.getWorkbookData()) {
            parser.parse(new InputSource(workbook));
        }
        boolean use1904Dates = date1904[0];
        boolean gregorian = defaultCalendarIsGregorian();
        return withUtc(() -> new DataFormatter(locale) {
            @Override
            public String formatRawCellContents(double value, int formatIndex, String formatString) {
                if (!use1904Dates && value >= 60 && value < 61 && formatString != null
                        && DateUtil.isADateFormat(formatIndex, formatString) && hasCalendarFields(formatString)) {
                    String date = "1900-02-29";
                    if (value == 60) return date;
                    long timeMillis = Math.min(86_399_999, Math.round((value - 60) * 86_400_000));
                    return date + "T" + LocalTime.ofNanoOfDay(timeMillis * 1_000_000).format(DateTimeFormatter.ISO_LOCAL_TIME);
                }
                if (DateUtil.isADateFormat(formatIndex, formatString)) {
                    // POI formats dates with the JVM's default calendar. Where that is not Gregorian (th-TH is Buddhist)
                    // the text would carry the wrong year, so the date comes from the serial number in ISO form instead.
                    if (!gregorian && hasCalendarFields(formatString)) return isoDateText(value, use1904Dates, formatString);
                    return asciiDigits(withUtc(() -> super.formatRawCellContents(value, formatIndex, formatString, use1904Dates)));
                }
                return super.formatRawCellContents(value, formatIndex, formatString, use1904Dates);
            }
        });
    }

    private static boolean defaultCalendarIsGregorian() {
        return "gregory".equals(Calendar.getInstance(Locale.getDefault(Locale.Category.FORMAT)).getCalendarType());
    }

    /**
     * A date is converted back from its text, and the converters read ASCII digits. POI formats dates with the digits of
     * the JVM's default locale (Arabic-Indic, Persian, Bengali, ...), which would show 2024-01-01 as '٢٠٢٤-٠١-٠١'.
     * Other characters, such as month names, are kept.
     */
    static String asciiDigits(String text) {
        StringBuilder out = null;
        for (int i = 0; i < text.length(); i++) {
            char c = text.charAt(i);
            if (c > 127 && Character.getType(c) == Character.DECIMAL_DIGIT_NUMBER) {
                if (out == null) out = new StringBuilder(text.length()).append(text, 0, i);
                out.append((char) ('0' + Character.digit(c, 10)));
            } else if (out != null) {
                out.append(c);
            }
        }
        return out == null ? text : out.toString();
    }

    /**
     * The date of a serial number as {@code yyyy-MM-dd}, with {@code HH:mm:ss} and the fraction the format asks for
     * when it also shows the time, whatever the default calendar.
     */
    static String isoDateText(double value, boolean use1904Dates, String format) {
        String fields = FORMAT_LITERAL.matcher(format.toLowerCase(Locale.ROOT)).replaceAll("");
        fields = FORMAT_MODIFIER.matcher(fields).replaceAll("");
        boolean time = fields.indexOf('h') >= 0 || fields.indexOf('s') >= 0;
        int fraction = 0;
        int dot = fields.indexOf("s.0");
        if (dot >= 0) {
            for (int i = dot + 2; i < fields.length() && fields.charAt(i) == '0' && fraction < 9; i++) fraction++;
        }
        LocalDateTime moment = DateUtil.getLocalDateTime(value, use1904Dates);
        // round to what is shown, as Excel does, and carry into the next second or day when that overflows
        long unit = (long) Math.pow(10, 9 - fraction);
        moment = moment.withNano(0).plusNanos((moment.getNano() + unit / 2) / unit * unit);
        if (!time) return DateTimeFormatter.ofPattern("yyyy-MM-dd", Locale.ROOT).format(moment);
        String pattern = "yyyy-MM-dd HH:mm:ss" + (fraction > 0 ? "." + "S".repeat(fraction) : "");
        return DateTimeFormatter.ofPattern(pattern, Locale.ROOT).format(moment);
    }

    private static <T> T withUtc(Supplier<T> formatting) {
        var previous = LocaleUtil.getUserTimeZone();
        LocaleUtil.setUserTimeZone(LocaleUtil.TIMEZONE_UTC);
        try {
            return formatting.get();
        } finally {
            LocaleUtil.setUserTimeZone(previous);
        }
    }

    private static boolean hasCalendarFields(String format) {
        String fields = FORMAT_LITERAL.matcher(format.toLowerCase(Locale.ROOT)).replaceAll("");
        boolean elapsed = ELAPSED_TIME.matcher(fields).find();
        fields = FORMAT_MODIFIER.matcher(fields).replaceAll("");
        return fields.indexOf('y') >= 0 || fields.indexOf('d') >= 0 || fields.contains("mmm")
                || (fields.indexOf('m') >= 0 && fields.indexOf('h') < 0 && fields.indexOf('s') < 0 && !elapsed);
    }

    /** Keeps worksheet value handling consistent on POI versions that predate ISO date support. */
    static class SheetXmlHandler extends XSSFSheetXMLHandler {
        static final String SPREADSHEET_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

        private final StringBuilder cachedText = new StringBuilder();
        private boolean stringCell;
        private boolean readingCachedText;
        private int rowNumber;
        private @Nullable String cellReference;

        SheetXmlHandler(@Nullable Styles styles, SharedStrings strings, SheetContentsHandler handler,
                        DataFormatter formatter) {
            super(styles, strings, handler, formatter, false);
        }

        @Override
        public void startElement(String uri, String localName, String qName, Attributes attributes)
                throws SAXException {
            if (SPREADSHEET_NAMESPACE.equals(uri)) {
                if ("row".equals(localName)) {
                    String rowReference = attributes.getValue("r");
                    rowNumber = rowReference == null ? rowNumber + 1 : Integer.parseInt(rowReference);
                    cellReference = null;
                } else if ("c".equals(localName)) {
                    String type = attributes.getValue("t");
                    stringCell = "str".equals(type);
                    String explicitReference = attributes.getValue("r");
                    if (explicitReference == null || "d".equals(type)) {
                        var normalized = new AttributesImpl(attributes);
                        if ("d".equals(type)) normalized.setValue(normalized.getIndex("t"), "str");
                        if (explicitReference == null) {
                            int column = cellReference == null ? 0 : getColumnIndex(cellReference) + 1;
                            if (column >= EXCEL_MAX_COLUMNS) {
                                throw new ExcelReadException("Implicit cell exceeds Excel maximum column (XFD) at row "
                                        + rowNumber);
                            }
                            explicitReference = CellReference.convertNumToColString(column) + rowNumber;
                            normalized.addAttribute("", "r", "r", "CDATA", explicitReference);
                        }
                        attributes = normalized;
                    }
                    cellReference = explicitReference;
                } else if (stringCell && "v".equals(localName)) {
                    readingCachedText = true;
                    cachedText.setLength(0);
                }
            }
            super.startElement(uri, localName, qName, attributes);
        }

        @Nullable String cellReference() {
            return cellReference;
        }

        @Override
        public void characters(char[] chars, int start, int length) throws SAXException {
            if (readingCachedText) {
                cachedText.append(chars, start, length);
            } else {
                super.characters(chars, start, length);
            }
        }

        @Override
        public void endElement(String uri, String localName, String qName) throws SAXException {
            if (readingCachedText && "v".equals(localName) && SPREADSHEET_NAMESPACE.equals(uri)) {
                readingCachedText = false;
                String text = cachedText.toString();
                if (text.contains("_x")) text = new XSSFRichTextString(text).getString();
                char[] chars = text.toCharArray();
                super.characters(chars, 0, chars.length);
            }
            super.endElement(uri, localName, qName);
        }
    }

    /**
     * Converts an Excel cell reference (e.g., "C5", "AA12") to a zero-based column index.
     *
     * @param cellReference The Excel cell reference (e.g., "C5", "AA10")
     * @return The zero-based column index
     * @throws ExcelReadException if the column index exceeds the Excel maximum (XFD = 16,384)
     */
    static int getColumnIndex(String cellReference) {
        int colIdx = 0;
        for (char c : cellReference.toCharArray()) {
            if (!Character.isLetter(c)) break;
            colIdx = colIdx * 26 + (Character.toUpperCase(c) - 'A' + 1);
            if (colIdx > EXCEL_MAX_COLUMNS) {
                throw new ExcelReadException("Column index exceeds Excel maximum (XFD): " + cellReference);
            }
        }
        return colIdx - 1;
    }
}
