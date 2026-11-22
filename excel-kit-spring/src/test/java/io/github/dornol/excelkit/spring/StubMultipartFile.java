package io.github.dornol.excelkit.spring;

import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayInputStream;
import java.io.File;
import java.io.IOException;
import java.io.InputStream;

/** A multipart file whose stream is whatever the test needs it to be, including a failing one. */
final class StubMultipartFile implements MultipartFile {

    interface StreamSupplier {
        InputStream open() throws IOException;
    }

    private final String originalFilename;
    private final long size;
    private final StreamSupplier stream;

    StubMultipartFile(String originalFilename, long size, StreamSupplier stream) {
        this.originalFilename = originalFilename;
        this.size = size;
        this.stream = stream;
    }

    static StubMultipartFile of(String originalFilename, byte[] content) {
        return new StubMultipartFile(originalFilename, content.length, () -> new ByteArrayInputStream(content));
    }

    /** A file whose stream cannot be opened. */
    static StubMultipartFile unopenable(String originalFilename) {
        return new StubMultipartFile(originalFilename, 10, () -> {
            throw new IOException("disk gone");
        });
    }

    /** A file whose stream delivers its bytes and then fails when it is closed. */
    static StubMultipartFile failingOnClose(String originalFilename, byte[] content) {
        return new StubMultipartFile(originalFilename, content.length, () -> new ByteArrayInputStream(content) {
            @Override
            public void close() throws IOException {
                throw new IOException("close failed");
            }
        });
    }

    /** A file whose stream opens but fails on the first read. */
    static StubMultipartFile unreadable(String originalFilename) {
        return new StubMultipartFile(originalFilename, 10, () -> new InputStream() {
            @Override
            public int read() throws IOException {
                throw new IOException("read failed");
            }
        });
    }

    @Override
    public String getName() {
        return "file";
    }

    @Override
    public String getOriginalFilename() {
        return originalFilename;
    }

    @Override
    public String getContentType() {
        return "application/octet-stream";
    }

    @Override
    public boolean isEmpty() {
        return size == 0;
    }

    @Override
    public long getSize() {
        return size;
    }

    @Override
    public byte[] getBytes() throws IOException {
        try (InputStream in = stream.open()) {
            return in.readAllBytes();
        }
    }

    @Override
    public InputStream getInputStream() throws IOException {
        return stream.open();
    }

    @Override
    public void transferTo(File dest) throws IOException, IllegalStateException {
        throw new UnsupportedOperationException();
    }
}
