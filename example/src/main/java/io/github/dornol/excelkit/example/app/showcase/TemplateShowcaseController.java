package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.common.DownloadFileType;
import io.github.dornol.excelkit.example.app.common.DownloadUtil;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.excel.ExcelTemplateWriter;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.bind.annotation.RequestMapping;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.io.ByteArrayInputStream;

@Controller
@RequestMapping("/showcase")
public class TemplateShowcaseController {

    // ============================================================
    // Template-based writing
    // ============================================================

    /**
     * Creates a template in memory (simulates loading a .xlsx template file),
     * fills in cell values and list data, and streams the result.
     */
    @GetMapping("/template-write")
    public ResponseEntity<StreamingResponseBody> templateWrite() {
        return DownloadUtil.builder("invoice", DownloadFileType.EXCEL).body(out -> {
            // 1. Create a template in memory (in real apps, load from classpath/filesystem)
            byte[] templateBytes;
            try (var twb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
                var sheet = twb.createSheet("Invoice");
                var titleRow = sheet.createRow(0);
                titleRow.createCell(0).setCellValue("INVOICE");
                sheet.addMergedRegion(new org.apache.poi.ss.util.CellRangeAddress(0, 0, 0, 3));

                sheet.createRow(2).createCell(0).setCellValue("Client:");
                sheet.createRow(3).createCell(0).setCellValue("Date:");

                var headerRow = sheet.createRow(5);
                headerRow.createCell(0).setCellValue("Product");
                headerRow.createCell(1).setCellValue("Qty");
                headerRow.createCell(2).setCellValue("Price");
                headerRow.createCell(3).setCellValue("Amount");

                var bos = new java.io.ByteArrayOutputStream();
                twb.write(bos);
                templateBytes = bos.toByteArray();
            }

            // 2. Fill template with data
            var products = ShowcaseData.sampleProducts();
            try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(templateBytes))) {
                // Cell-level writes (fill placeholders)
                writer.cell("B3", "Acme Corporation")
                      .cell("B4", java.time.LocalDate.now());

                // List streaming (starting at row 6, headers are already in row 5)
                writer.<ProductDto>list(6)
                      .column("Product", ProductDto::name)
                      .column("Qty", p -> 10, c -> c.type(ExcelDataType.INTEGER))
                      .column("Price", ProductDto::price, c -> c.type(ExcelDataType.INTEGER))
                      .column("Amount", p -> p.price() * 10, c -> c.type(ExcelDataType.INTEGER))
                      .afterData(ctx -> {
                          var row = ctx.getSheet().createRow(ctx.getCurrentRow());
                          row.createCell(0).setCellValue("Total");
                          row.createCell(3).setCellFormula(
                                  "SUM(D7:D" + ctx.getCurrentRow() + ")");
                          return ctx.getCurrentRow() + 1;
                      })
                      .write(products.stream());

                writer.finish().consumeOutputStream(out);
            }
        });
    }

    /**
     * Demonstrates cell-only template writing (e.g., certificate/document style).
     */
    @GetMapping("/template-cell-only")
    public ResponseEntity<StreamingResponseBody> templateCellOnly() {
        return DownloadUtil.builder("certificate", DownloadFileType.EXCEL).body(out -> {
            byte[] templateBytes;
            try (var twb = new org.apache.poi.xssf.usermodel.XSSFWorkbook()) {
                var sheet = twb.createSheet("Certificate");
                sheet.createRow(1).createCell(1).setCellValue("Certificate of Employment");
                sheet.createRow(3).createCell(0).setCellValue("Name:");
                sheet.createRow(4).createCell(0).setCellValue("Department:");
                sheet.createRow(5).createCell(0).setCellValue("Position:");
                sheet.createRow(6).createCell(0).setCellValue("Join Date:");
                sheet.createRow(8).createCell(0).setCellValue("Issue Date:");

                var bos = new java.io.ByteArrayOutputStream();
                twb.write(bos);
                templateBytes = bos.toByteArray();
            }

            try (var writer = new ExcelTemplateWriter(new ByteArrayInputStream(templateBytes))) {
                writer.cell("B4", "John Doe")
                      .cell("B5", "Engineering")
                      .cell("B6", "Senior Developer")
                      .cell("B7", java.time.LocalDate.of(2020, 3, 15))
                      .cell("B9", java.time.LocalDate.now())
                      .finish()
                      .consumeOutputStream(out);
            }
        });
    }
}
