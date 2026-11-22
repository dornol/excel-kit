package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelImage;
import com.sun.net.httpserver.HttpServer;
import org.apache.poi.ss.usermodel.Workbook;
import org.junit.jupiter.api.AfterEach;
import org.junit.jupiter.api.BeforeEach;
import org.junit.jupiter.api.Test;

import java.io.IOException;
import java.io.UncheckedIOException;
import java.net.InetSocketAddress;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertDoesNotThrow;
import static org.junit.jupiter.api.Assertions.assertThrows;

/** How a downloaded image gets its type and its size limit, against a local HTTP server. */
class ImageTypeDetectionTest {

    private static final byte[] JPEG = {(byte) 0xFF, (byte) 0xD8, (byte) 0xFF, (byte) 0xE0, 0, 16, 'J', 'F', 'I', 'F'};
    private static final byte[] PNG = {(byte) 0x89, 'P', 'N', 'G', 13, 10, 26, 10, 0, 0};
    private static final byte[] SINGLE_FF = {(byte) 0xFF, 0x01, 0x02};

    private HttpServer server;

    @BeforeEach
    void start() throws IOException {
        server = HttpServer.create(new InetSocketAddress("127.0.0.1", 0), 0);
        server.createContext("/jpeg", exchange -> reply(exchange, JPEG, true));
        server.createContext("/png", exchange -> reply(exchange, PNG, true));
        server.createContext("/ff", exchange -> reply(exchange, SINGLE_FF, true));
        server.createContext("/photo.jpg", exchange -> reply(exchange, PNG, true));
        server.createContext("/PHOTO.JPEG", exchange -> reply(exchange, PNG, true));
        server.createContext("/drawing.png", exchange -> reply(exchange, JPEG, true));
        server.createContext("/chunked", exchange -> reply(exchange, new byte[20], false));
        server.createContext("/sized", exchange -> reply(exchange, new byte[20], true));
        server.start();
    }

    @AfterEach
    void stop() {
        server.stop(0);
    }

    private static void reply(com.sun.net.httpserver.HttpExchange exchange, byte[] body, boolean knownLength) throws IOException {
        exchange.sendResponseHeaders(200, knownLength ? body.length : 0); // 0 means chunked
        try (var out = exchange.getResponseBody()) {
            out.write(body);
        }
    }

    private String url(String path) {
        return "http://127.0.0.1:" + server.getAddress().getPort() + path;
    }

    @Test
    void theExtensionOfTheUrlDecidesFirst() {
        // the bytes say PNG, the name says JPEG (in either case and spelling), and the name wins
        assertEquals(Workbook.PICTURE_TYPE_JPEG, ExcelImage.fromUrl(url("/photo.jpg")).imageType());
        assertEquals(Workbook.PICTURE_TYPE_JPEG, ExcelImage.fromUrl(url("/PHOTO.JPEG")).imageType());
        // and the other way round
        assertEquals(Workbook.PICTURE_TYPE_PNG, ExcelImage.fromUrl(url("/drawing.png")).imageType());
    }

    @Test
    void aJpegIsRecognisedByItsFirstTwoBytes() {
        assertEquals(Workbook.PICTURE_TYPE_JPEG, ExcelImage.fromUrl(url("/jpeg")).imageType());
    }

    @Test
    void anythingElseIsTakenForAPng() {
        assertEquals(Workbook.PICTURE_TYPE_PNG, ExcelImage.fromUrl(url("/png")).imageType());
        // 0xFF alone is not a JPEG signature
        assertEquals(Workbook.PICTURE_TYPE_PNG, ExcelImage.fromUrl(url("/ff")).imageType());
    }

    @Test
    void theDownloadLimitIsInclusive() {
        assertDoesNotThrow(() -> ExcelImage.fromUrl(url("/sized"), 20));
        assertThrows(UncheckedIOException.class, () -> ExcelImage.fromUrl(url("/sized"), 19));
        // a server that does not say how long the body is cannot get past the limit either
        assertDoesNotThrow(() -> ExcelImage.fromUrl(url("/chunked"), 20));
        assertThrows(UncheckedIOException.class, () -> ExcelImage.fromUrl(url("/chunked"), 19));
    }

    @Test
    void theLimitItselfMustBeSensible() {
        assertThrows(IllegalArgumentException.class, () -> ExcelImage.fromUrl(url("/jpeg"), 0));
        assertThrows(IllegalArgumentException.class, () -> ExcelImage.fromUrl(url("/jpeg"), -1));
        assertThrows(IllegalArgumentException.class, () -> ExcelImage.fromUrl(url("/jpeg"), Integer.MAX_VALUE));
        assertDoesNotThrow(() -> ExcelImage.fromUrl(url("/jpeg"), JPEG.length));
    }

    @Test
    void anImageSpansAtLeastOneColumnAndOneRow() {
        ExcelImage image = ExcelImage.png(PNG);
        assertEquals(1, image.width());
        assertEquals(1, image.height());
        assertDoesNotThrow(() -> image.size(1, 1));
        assertThrows(IllegalArgumentException.class, () -> image.size(0, 1));
        assertThrows(IllegalArgumentException.class, () -> image.size(1, 0));
        image.size(3, 4);
        assertEquals(3, image.width());
        assertEquals(4, image.height());
    }
}
