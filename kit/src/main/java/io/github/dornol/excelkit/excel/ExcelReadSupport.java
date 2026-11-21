package io.github.dornol.excelkit.excel;

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

import java.time.LocalTime;
import java.time.format.DateTimeFormatter;
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
    static void decrypt(java.nio.file.Path encrypted, String password, java.nio.file.Path target) throws Exception {
        try (org.apache.poi.poifs.filesystem.POIFSFileSystem fs =
                     new org.apache.poi.poifs.filesystem.POIFSFileSystem(encrypted.toFile(), true)) {
            org.apache.poi.poifs.crypt.EncryptionInfo info = new org.apache.poi.poifs.crypt.EncryptionInfo(fs);
            org.apache.poi.poifs.crypt.Decryptor decryptor = org.apache.poi.poifs.crypt.Decryptor.getInstance(info);
            if (!decryptor.verifyPassword(password)) {
                throw new ExcelReadException("Invalid password for encrypted Excel file");
            }
            try (java.io.InputStream decrypted = decryptor.getDataStream(fs);
                 java.io.OutputStream out = java.nio.file.Files.newOutputStream(target)) {
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

    /** Uses the workbook's date system for numeric cells and cached formula results. */
    static DataFormatter dataFormatter(XSSFReader reader) throws Exception {
        return dataFormatter(reader, LocaleUtil.getUserLocale());
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
                    return withUtc(() -> super.formatRawCellContents(value, formatIndex, formatString, use1904Dates));
                }
                return super.formatRawCellContents(value, formatIndex, formatString, use1904Dates);
            }
        });
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
