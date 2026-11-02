package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.util.XMLHelper;
import org.apache.poi.xssf.eventusermodel.XSSFReader;
import org.xml.sax.InputSource;
import org.xml.sax.Attributes;
import org.xml.sax.helpers.DefaultHandler;

/**
 * Package-private utility methods shared by {@link ExcelReadHandler} and {@link ExcelReader}
 * to eliminate duplicate read logic.
 *
 * @author dhkim
 */
class ExcelReadSupport {

    /** Excel maximum column count (XFD = 16,384). */
    static final int EXCEL_MAX_COLUMNS = 16_384;

    private ExcelReadSupport() {
    }

    /** Uses the workbook's date system for numeric cells and cached formula results. */
    static DataFormatter dataFormatter(XSSFReader reader) throws Exception {
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
        return new DataFormatter() {
            @Override
            public String formatRawCellContents(double value, int formatIndex, String formatString) {
                return super.formatRawCellContents(value, formatIndex, formatString, use1904Dates);
            }
        };
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
