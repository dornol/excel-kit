package io.github.dornol.excelkit.excel;

import io.github.dornol.excelkit.excel.feature.ExcelImage;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.junit.jupiter.api.Test;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.*;
import java.util.stream.Stream;

import static org.junit.jupiter.api.Assertions.*;

class ImageEmbeddingTest {
    // Minimal valid 1x1 PNG
    private static final byte[] TINY_PNG = {
            (byte) 0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
            0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
            0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
            0x08, 0x02, 0x00, 0x00, 0x00, (byte) 0x90, 0x77, 0x53,
            (byte) 0xDE, 0x00, 0x00, 0x00, 0x0C, 0x49, 0x44, 0x41,
            0x54, 0x08, (byte) 0xD7, 0x63, (byte) 0xF8, (byte) 0xCF,
            (byte) 0xC0, 0x00, 0x00, 0x00, 0x02, 0x00, 0x01,
            (byte) 0xE2, 0x21, (byte) 0xBC, 0x33, 0x00, 0x00, 0x00,
            0x00, 0x49, 0x45, 0x4E, 0x44, (byte) 0xAE, 0x42, 0x60,
            (byte) 0x82
    };

    @Test
    void image_multipleImages() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("Name", s -> s)
                .column("Pic", s -> ExcelImage.png(TINY_PNG), c -> c.type(ExcelDataType.IMAGE))
                .write(Stream.of("A", "B", "C"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals(3, wb.getAllPictures().size());
        }
    }

    @Test
    void image_withTextColumns() throws IOException {
        ByteArrayOutputStream out = new ByteArrayOutputStream();
        ExcelWriter.<String>create()
                .column("ID", s -> s)
                .column("Name", s -> "Item-" + s)
                .column("Photo", s -> ExcelImage.png(TINY_PNG), c -> c.type(ExcelDataType.IMAGE))
                .column("Notes", s -> "Note-" + s)
                .write(Stream.of("1"))
                .writeTo(out);

        try (var wb = new XSSFWorkbook(new ByteArrayInputStream(out.toByteArray()))) {
            assertEquals("1", wb.getSheetAt(0).getRow(1).getCell(0).getStringCellValue());
            assertEquals("Item-1", wb.getSheetAt(0).getRow(1).getCell(1).getStringCellValue());
            assertEquals("Note-1", wb.getSheetAt(0).getRow(1).getCell(3).getStringCellValue());
            assertEquals(1, wb.getAllPictures().size());
        }
    }

    @Test
    void excelImage_record_equality() {
        byte[] data = {1, 2, 3};
        ExcelImage img1 = ExcelImage.png(data);
        ExcelImage img2 = new ExcelImage(data, Workbook.PICTURE_TYPE_PNG);
        assertEquals(img1.imageType(), img2.imageType());
    }
}
