package io.github.dornol.excelkit.core;

import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.nio.charset.StandardCharsets;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertNull;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** The decisions of {@link TabularFileDetector} at their edges. */
class TabularFileDetectorDetailsTest {

    private static TabularDetectionResult detect(byte[] bytes) {
        return TabularFileDetector.detectDetailed(new ByteArrayInputStream(bytes));
    }

    private static TabularDetectionResult detect(String text) {
        return detect(text.getBytes(StandardCharsets.UTF_8));
    }

    private static Character delimiter(String text) {
        return detect(text).delimiter();
    }

    // ---- signatures

    @Test
    void anXlsxStartsWithAFourByteSignature() {
        assertEquals(TabularFileType.XLSX, detect(new byte[]{0x50, 0x4b, 0x03, 0x04}).type());
        assertEquals(DetectionConfidence.HIGH, detect(new byte[]{0x50, 0x4b, 0x03, 0x04}).confidence());
        // three bytes of it, or four that differ, are not a workbook (their control bytes make them binary)
        assertEquals(TabularFileType.UNKNOWN, detect(new byte[]{0x50, 0x4b, 0x03}).type());
        assertEquals(TabularFileType.UNKNOWN, detect(new byte[]{0x50, 0x4b, 0x03, 0x05}).type());
        // and the signature must be at the start, not just somewhere in the sample
        assertEquals(TabularFileType.UNKNOWN, detect(new byte[]{'x', 0x50, 0x4b, 0x03, 0x04}).type());
    }

    @Test
    void anXlsStartsWithAnEightByteSignature() {
        byte[] xls = {(byte) 0xd0, (byte) 0xcf, 0x11, (byte) 0xe0, (byte) 0xa1, (byte) 0xb1, 0x1a, (byte) 0xe1};
        assertEquals(TabularFileType.XLS, detect(xls).type());
        assertEquals(DetectionConfidence.HIGH, detect(xls).confidence());
        assertEquals(TabularFileType.UNKNOWN, detect(java.util.Arrays.copyOf(xls, 7)).type());
    }

    // ---- is it text

    @Test
    void aNulByteMeansBinary() {
        assertEquals(TabularFileType.UNKNOWN, detect("a,b\u0000c").type());
        assertEquals(DetectionConfidence.LOW, detect("a,b\u0000c").confidence());
    }

    @Test
    void textIsAnythingBelowFivePercentControlCharacters() {
        // one control character in 21 characters (4.8%) is text, one in 20 (exactly 5%) is binary
        assertEquals(TabularFileType.CSV, detect("a".repeat(20) + "\u0001").type());
        assertEquals(TabularFileType.UNKNOWN, detect("a".repeat(19) + "\u0001").type());
        // two in 40 is exactly 5% too
        assertEquals(TabularFileType.UNKNOWN, detect("a".repeat(19) + "\u0001" + "a".repeat(19) + "\u0002").type());
        assertEquals(TabularFileType.CSV, detect("a".repeat(20) + "\u0001" + "a".repeat(19) + "\u0002").type());
    }

    @Test
    void tabsAndLineBreaksAreNotControlCharacters() {
        assertEquals(TabularFileType.CSV, detect("a\tb\r\nc\td\r\n").type());
    }

    // ---- the delimiter

    @Test
    void theMostFrequentCandidateWinsAndTiesGoToTheEarlierOne() {
        assertEquals(',', delimiter("a,b,c;d"));
        assertEquals(';', delimiter("a;b;c,d"));
        assertEquals(',', delimiter("a,b;c"));      // tie: comma comes first
        assertEquals('\t', delimiter("a\tb;c"));    // tie: tab before semicolon
        assertEquals(';', delimiter("a;b|c"));      // tie: semicolon before pipe
        assertEquals('|', delimiter("a|b|c;d"));
    }

    @Test
    void noDelimiterMeansLowConfidenceAndNoDelimiter() {
        TabularDetectionResult single = detect("just one column\n");
        assertEquals(TabularFileType.CSV, single.type());
        assertEquals(DetectionConfidence.LOW, single.confidence());
        assertNull(single.delimiter());

        TabularDetectionResult split = detect("a,b\n");
        assertEquals(DetectionConfidence.MEDIUM, split.confidence());
        assertEquals(',', split.delimiter());
    }

    @Test
    void delimitersInsideQuotesAreNotCounted() {
        assertEquals(';', delimiter("\"a,b,c,d\";x;y"));
        // a doubled quote inside a quoted field does not end it
        assertEquals(';', delimiter("\"a\"\",b,\"\",c\";x"));
        // a quote that does not start a field is an ordinary character
        assertEquals(',', delimiter("ab\"c,d,e;f"));
    }

    @Test
    void spacesBeforeAQuoteKeepTheFieldQuoted() {
        assertEquals(';', delimiter(" \"x,y,z\";a"));
    }

    @Test
    void onlyTheHeaderRecordCounts() {
        assertEquals(',', delimiter("a,b\nc;d;e;f\n"));
        assertEquals(',', delimiter("a,b\r\nc;d;e;f\r\n"));
        assertEquals(',', delimiter("a,b\rc;d;e;f\r"));
    }

    @Test
    void theHeaderRecordCanBeALaterOne() {
        byte[] data = "junk\nmore junk\nx;y;z\nrow,with,commas\n".getBytes(StandardCharsets.UTF_8);
        assertEquals(';', TabularFileDetector.detectDetailed(new ByteArrayInputStream(data), 8192, null, '"', '\\', 2).delimiter());
        assertEquals(';', TabularFileDetector.detectDetailed(
                new ByteArrayInputStream("a,b\r\nx;y;z\r\n".getBytes(StandardCharsets.UTF_8)), 8192, null, '"', '\\', 1).delimiter());
    }

    @Test
    void aHeaderBeyondTheSampleIsReported() {
        byte[] data = "a,b\nc;d\n".getBytes(StandardCharsets.UTF_8);
        assertThrows(ExcelKitException.class, () -> TabularFileDetector.detectDetailed(
                new ByteArrayInputStream(data), 64, null, '"', '\\', 5));
        // a header record that starts inside the sample but is cut off by it still decides the delimiter
        byte[] cutOff = ("junk\n" + "x;y;" + "z".repeat(100)).getBytes(StandardCharsets.UTF_8);
        assertEquals(';', TabularFileDetector.detectDetailed(
                new ByteArrayInputStream(cutOff), 64, null, '"', '\\', 1).delimiter());
        // one that would start exactly where the sample ends does not exist
        byte[] atTheEnd = ("a".repeat(63) + "\n" + "x;y").getBytes(StandardCharsets.UTF_8);
        assertThrows(ExcelKitException.class, () -> TabularFileDetector.detectDetailed(
                new ByteArrayInputStream(atTheEnd), 64, null, '"', '\\', 1));
    }

    @Test
    void anEscapeCharacterHidesTheDelimiterAfterIt() {
        // with \ as the escape character, "a\,b" is one field in which the comma is escaped
        assertEquals(';', TabularFileDetector.detectDetailed(new ByteArrayInputStream(
                "a\\,b\\,c;d;e".getBytes(StandardCharsets.UTF_8)), 8192, null, '"', '\\').delimiter());
        // without one, the commas count
        assertEquals(',', TabularFileDetector.detectDetailed(new ByteArrayInputStream(
                "a\\,b\\,c;d".getBytes(StandardCharsets.UTF_8)), 8192, null, '"', '\0').delimiter());
    }

    // ---- the encoding

    @Test
    void aByteOrderMarkSelectsTheCharsetAndDoesNotCountAsData() {
        byte[] utf16le = "a,b\n".getBytes(StandardCharsets.UTF_16LE);
        byte[] withBom = new byte[utf16le.length + 2];
        withBom[0] = (byte) 0xff;
        withBom[1] = (byte) 0xfe;
        System.arraycopy(utf16le, 0, withBom, 2, utf16le.length);
        TabularDetectionResult result = detect(withBom);
        assertEquals(TabularFileType.CSV, result.type());
        assertEquals(StandardCharsets.UTF_16LE, result.charset());
        assertEquals(',', result.delimiter());

        byte[] utf16be = "a;b\n".getBytes(StandardCharsets.UTF_16BE);
        byte[] bigEndian = new byte[utf16be.length + 2];
        bigEndian[0] = (byte) 0xfe;
        bigEndian[1] = (byte) 0xff;
        System.arraycopy(utf16be, 0, bigEndian, 2, utf16be.length);
        assertEquals(StandardCharsets.UTF_16BE, detect(bigEndian).charset());
        assertEquals(';', detect(bigEndian).delimiter());

        assertEquals(StandardCharsets.UTF_8, detect("a,b").charset());
        assertEquals(',', delimiter("﻿a,b"));
    }

    @Test
    void aConfiguredCharsetBeatsTheOneInferredFromTheSample() {
        byte[] latin1 = "café;x\n".getBytes(StandardCharsets.ISO_8859_1);
        TabularDetectionResult result = TabularFileDetector.detectDetailed(
                new ByteArrayInputStream(latin1), 8192, StandardCharsets.ISO_8859_1, '"', '\\');
        assertEquals(StandardCharsets.ISO_8859_1, result.charset());
        assertEquals(';', result.delimiter());
    }

    @Test
    void theSampleSizeIsBoundedBelow() {
        assertThrows(IllegalArgumentException.class, () -> TabularFileDetector.detectDetailed(
                new ByteArrayInputStream(new byte[]{'a'}), 63));
        assertEquals(TabularFileType.CSV, TabularFileDetector.detectDetailed(
                new ByteArrayInputStream("a,b".getBytes(StandardCharsets.UTF_8)), 64).type());
    }

    @Test
    void anInputThatCannotRewindIsRefused() {
        assertThrows(IllegalArgumentException.class, () -> TabularFileDetector.detectDetailed(
                new java.io.FilterInputStream(new ByteArrayInputStream(new byte[]{'a'})) {
                    @Override public boolean markSupported() { return false; }
                }));
    }
}
