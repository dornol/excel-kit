package io.github.dornol.excelkit.excel.read;

import io.github.dornol.excelkit.core.read.ReadSecurityException;
import io.github.dornol.excelkit.core.read.ReadSecurityPolicy;
import javax.xml.stream.XMLInputFactory;
import javax.xml.stream.XMLStreamConstants;
import javax.xml.stream.XMLStreamException;
import javax.xml.stream.XMLStreamReader;
import java.io.FilterInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.nio.file.Path;
import java.util.Locale;
import java.util.zip.ZipEntry;
import java.util.zip.ZipFile;

/** Bounded preflight scanner for formula and external-link policy enforcement. */
final class ExcelSecurityScanner {
    private static final String MAIN_NAMESPACE = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
    private static final String STRICT_NAMESPACE = "http://purl.oclc.org/ooxml/spreadsheetml/main";
    private static final XMLInputFactory XML = xmlFactory();

    private ExcelSecurityScanner() {}

    private static XMLInputFactory xmlFactory() {
        XMLInputFactory factory = XMLInputFactory.newFactory();
        factory.setProperty(XMLInputFactory.IS_NAMESPACE_AWARE, true);
        factory.setProperty(XMLInputFactory.SUPPORT_DTD, false);
        factory.setProperty(XMLInputFactory.IS_SUPPORTING_EXTERNAL_ENTITIES, false);
        return factory;
    }

    static void scan(Path file, ReadSecurityPolicy policy) throws IOException {
        long total = 0;
        int entryCount = 0;
        try (ZipFile zip = new ZipFile(file.toFile())) {
            var zipEntries = zip.entries();
            while (zipEntries.hasMoreElements()) {
                ZipEntry entry = zipEntries.nextElement();
                if (++entryCount > policy.maxZipEntries()) {
                    throw security(ReadSecurityException.Reason.ZIP_ENTRIES,
                            "Workbook exceeds ZIP entry limit: " + policy.maxZipEntries());
                }
                // OPC part names are case-insensitive and a worksheet part may live at any path.
                String name = entry.getName().toLowerCase(Locale.ROOT);
                if (!policy.allowExternalLinks()) {
                    if (name.startsWith("xl/externallinks/")) {
                        throw security(ReadSecurityException.Reason.EXTERNAL_LINK,
                                "External workbook links are not allowed");
                    }
                    if (name.endsWith(".rels")) {
                        validateEntry(entry, policy);
                        try (InputStream input = zip.getInputStream(entry)) {
                            if (containsExternalWorkbookLink(input, policy.maxScannedEntryBytes())) {
                                throw security(ReadSecurityException.Reason.EXTERNAL_LINK,
                                        "External workbook links are not allowed");
                            }
                        }
                    }
                }
                if (!policy.allowFormulas()) {
                    boolean standardPath = name.startsWith("xl/worksheets/") && name.endsWith(".xml");
                    ScanResult result = scanWorksheet(zip, entry, standardPath, policy);
                    total += result.bytes();
                    if (total > policy.maxTotalScannedBytes())
                        throw security(ReadSecurityException.Reason.TOTAL_SCAN_SIZE,
                                "Workbook security scan exceeds total byte limit");
                    if (result.formula()) throw security(ReadSecurityException.Reason.FORMULA,
                            "Excel formulas are not allowed");
                }
            }
        }
    }

    private static void validateEntry(ZipEntry entry, ReadSecurityPolicy policy) {
        long size = entry.getSize(), compressed = entry.getCompressedSize();
        if (size > policy.maxScannedEntryBytes()) throw security(ReadSecurityException.Reason.ENTRY_SIZE,
                "Worksheet XML exceeds security scan entry limit: " + size);
        if (size > 0 && compressed > 0 && (double) size / compressed > policy.maxCompressionRatio())
            throw security(ReadSecurityException.Reason.COMPRESSION_RATIO,
                    "Worksheet XML exceeds compression ratio limit");
    }

    /**
     * Looks for formula elements in a worksheet part. A part is a worksheet when its root element is
     * the SpreadsheetML {@code worksheet}, wherever it lives in the package and whichever namespace
     * prefix it uses; other parts (charts, drawings, media) are not scanned.
     */
    private static ScanResult scanWorksheet(ZipFile zip, ZipEntry entry, boolean standardPath,
                                            ReadSecurityPolicy policy) throws IOException {
        if (standardPath) validateEntry(entry, policy);
        try (LimitedInput input = new LimitedInput(zip.getInputStream(entry))) {
            XMLStreamReader reader = null;
            boolean worksheet = standardPath;
            if (standardPath) input.limit = policy.maxScannedEntryBytes();
            try {
                reader = XML.createXMLStreamReader(input);
                while (reader.hasNext()) {
                    if (reader.next() != XMLStreamConstants.START_ELEMENT) continue;
                    if (!worksheet) {
                        // Any other part is a worksheet only when its root element says so.
                        if (!isSpreadsheetElement(reader, "worksheet", false)) return new ScanResult(false, 0);
                        worksheet = true;
                        validateEntry(entry, policy);
                        input.limit = policy.maxScannedEntryBytes();
                    } else if (isSpreadsheetElement(reader, "f", standardPath)) {
                        return new ScanResult(true, input.count);
                    }
                }
                return new ScanResult(false, worksheet ? input.count : 0);
            } catch (XMLStreamException | RuntimeException e) {
                if (e instanceof ReadSecurityException security) throw security;
                // Not XML (media, binaries) is not a worksheet; an unreadable worksheet fails closed.
                if (worksheet) {
                    throw new ReadSecurityException(ReadSecurityException.Reason.FORMULA,
                            "Worksheet XML could not be scanned for formulas", e);
                }
                return new ScanResult(false, 0);
            } finally {
                if (reader != null) {
                    try { reader.close(); } catch (XMLStreamException ignored) { /* best effort */ }
                }
            }
        }
    }

    private static boolean isSpreadsheetElement(XMLStreamReader reader, String localName, boolean allowNoNamespace) {
        String namespace = reader.getNamespaceURI();
        return localName.equals(reader.getLocalName())
                && (MAIN_NAMESPACE.equals(namespace) || STRICT_NAMESPACE.equals(namespace)
                || (allowNoNamespace && (namespace == null || namespace.isEmpty())));
    }

    /** Counts bytes and enforces the entry limit once the part is known to be a worksheet. */
    private static final class LimitedInput extends FilterInputStream {
        long count;
        long limit = Long.MAX_VALUE;

        LimitedInput(InputStream in) { super(in); }

        @Override public int read() throws IOException {
            int value = super.read();
            if (value >= 0) add(1);
            return value;
        }

        @Override public int read(byte[] buffer, int offset, int length) throws IOException {
            int read = super.read(buffer, offset, length);
            if (read > 0) add(read);
            return read;
        }

        private void add(int bytes) {
            count += bytes;
            if (count > limit) throw security(ReadSecurityException.Reason.ENTRY_SIZE,
                    "Worksheet XML exceeds security scan entry limit");
        }
    }

    private static boolean containsExternalWorkbookLink(InputStream input, long maximum) throws IOException {
        final String needle = "externalLink";
        byte[] buffer = new byte[8192];
        // Only the tail that could be the start of a match survives from one chunk to the next.
        StringBuilder window = new StringBuilder(buffer.length + needle.length());
        long total = 0;
        for (int read; (read = input.read(buffer)) >= 0;) {
            total += read;
            if (total > maximum) {
                throw security(ReadSecurityException.Reason.ENTRY_SIZE,
                        "Relationship entry exceeds security scan entry limit");
            }
            for (int i = 0; i < read; i++) window.append((char) (buffer[i] & 0xFF));
            if (window.indexOf(needle) >= 0) return true;
            window.delete(0, Math.max(0, window.length() - (needle.length() - 1)));
        }
        return false;
    }

    private static ReadSecurityException security(ReadSecurityException.Reason reason, String message) {
        return new ReadSecurityException(reason, message);
    }
    private record ScanResult(boolean formula, long bytes) {}
}
