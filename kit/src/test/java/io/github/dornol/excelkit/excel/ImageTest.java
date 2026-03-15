package io.github.dornol.excelkit.excel;

import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFPictureData;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ImageTest {

    // Minimal valid 1x1 PNG
    private static final byte[] TINY_PNG = {
            (byte) 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A, // PNG signature
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52, // IHDR chunk
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01, // 1x1 dimensions
            0x08, 0x02, 0x00, 0x00, 0x00, (byte) 0x90, 0x77, 0x53,
            (byte) 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, (byte) 0xD7, 0x63, (byte) 0xF8, (byte) 0xCF,
            (byte) 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01,
            (byte) 0xE2, 0x21, (byte) 0xBC, 0x33, 0x00, 0x00, 0x00,
            0x00, 0x49, 0x45, 0x4E, 0x44, (byte) 0xAE, 0x42, 0x60,
            (byte) 0x82
    };

    @Test
    void image_shouldEmbedInExcel() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Name", s -> s)
                .addColumn("Photo", s -> ExcelImage.png(TINY_PNG), c -> c.type(ExcelDataType.IMAGE))
                .write(Stream.of("Alice"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            var pictures = wb.getAllPictures();
            assertFalse(pictures.isEmpty(), "Workbook should contain at least one picture");
            XSSFPictureData pic = (XSSFPictureData) pictures.get(0);
            assertEquals(Workbook.PICTURE_TYPE_PNG, pic.getPictureType());
        }
    }

    @Test
    void excelImage_png_factory() {
        ExcelImage img = ExcelImage.png(TINY_PNG);
        assertEquals(Workbook.PICTURE_TYPE_PNG, img.imageType());
        assertArrayEquals(TINY_PNG, img.data());
    }

    @Test
    void excelImage_jpeg_factory() {
        byte[] data = new byte[]{1, 2, 3};
        ExcelImage img = ExcelImage.jpeg(data);
        assertEquals(Workbook.PICTURE_TYPE_JPEG, img.imageType());
        assertArrayEquals(data, img.data());
    }

    @Test
    void image_nonImageValue_shouldFallbackToString() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        new ExcelWriter<String>()
                .addColumn("Data", s -> s, c -> c.type(ExcelDataType.IMAGE))
                .write(Stream.of("not-an-image"))
                .consumeOutputStream(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("not-an-image", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
        }
    }
}
