package io.github.dornol.excelkit.testsupport;

import org.apache.poi.xssf.streaming.SXSSFWorkbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.IOException;

/**
 * An {@link SXSSFWorkbook} whose {@code close()} also deletes its temp files.
 * <p>
 * Older Apache POI releases (for example 5.2.5) do not call {@code dispose()} from {@code close()},
 * so a test that builds a workbook directly would leave one temp file per sheet behind on them and
 * trip {@link TempFileLeakGuard}. Tests that go through the library need nothing special: it disposes
 * the workbook itself.
 */
public class DisposingSxssfWorkbook extends SXSSFWorkbook {

    public DisposingSxssfWorkbook() {
        super();
    }

    public DisposingSxssfWorkbook(int rowAccessWindowSize) {
        super(rowAccessWindowSize);
    }

    public DisposingSxssfWorkbook(XSSFWorkbook template, int rowAccessWindowSize) {
        super(template, rowAccessWindowSize);
    }

    @Override
    public void close() throws IOException {
        try {
            super.close();
        } finally {
            dispose();
        }
    }
}
