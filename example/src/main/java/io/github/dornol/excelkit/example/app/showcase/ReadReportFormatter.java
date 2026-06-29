package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductReadDto;
import io.github.dornol.excelkit.spring.CellErrorResponse;
import io.github.dornol.excelkit.spring.UploadError;
import io.github.dornol.excelkit.spring.UploadResult;

final class ReadReportFormatter {

    private ReadReportFormatter() {
    }

    static String formatTextReport(UploadResult<ProductReadDto> report) {
        StringBuilder sb = new StringBuilder();
        sb.append("=== Name-Based %s Read Result ===\n".formatted(report.type()));
        sb.append("Success: %d rows, Errors: %d rows\n\n".formatted(report.successCount(), report.errorCount()));
        report.rows().forEach(p -> sb.append(p).append("\n"));
        if (!report.errors().isEmpty()) {
            sb.append("\n--- Errors ---\n");
            report.errors().forEach(e -> sb.append(formatReadError(e)).append("\n"));
        }
        return sb.toString();
    }

    static String formatHtmlReport(UploadResult<ProductReadDto> report) {
        StringBuilder sb = new StringBuilder();
        sb.append("<!doctype html><html lang=\"ko\"><head><meta charset=\"utf-8\">")
                .append("<title>").append(escapeHtml(report.type())).append(" Read Result</title>")
                .append("<style>body{font-family:-apple-system,BlinkMacSystemFont,sans-serif;margin:24px;line-height:1.5}")
                .append("table{border-collapse:collapse;margin-top:12px}th,td{border:1px solid #ddd;padding:6px 8px}")
                .append("th{background:#f6f8fa}.error{color:#b00020}</style></head><body>");
        sb.append("<h1>Name-Based ").append(escapeHtml(report.type())).append(" Read Result</h1>");
        sb.append("<p>Success: ").append(report.successCount())
                .append(" rows, Errors: ").append(report.errorCount()).append(" rows</p>");
        if (!report.rows().isEmpty()) {
            sb.append("<h2>Rows</h2><table><thead><tr><th>Name</th><th>Category</th><th>Price</th>")
                    .append("<th>Quantity</th><th>Discount</th></tr></thead><tbody>");
            for (ProductReadDto row : report.rows()) {
                sb.append("<tr><td>").append(escapeHtml(row.getName()))
                        .append("</td><td>").append(escapeHtml(row.getCategory()))
                        .append("</td><td>").append(row.getPrice())
                        .append("</td><td>").append(row.getQuantity())
                        .append("</td><td>").append(row.getDiscount()).append("</td></tr>");
            }
            sb.append("</tbody></table>");
        }
        if (!report.errors().isEmpty()) {
            sb.append("<h2>Errors</h2><ul class=\"error\">");
            report.errors().forEach(e -> sb.append("<li>").append(escapeHtml(formatReadError(e))).append("</li>"));
            sb.append("</ul>");
        }
        return sb.append("</body></html>").toString();
    }

    private static String formatReadError(UploadError readError) {
        StringBuilder sb = new StringBuilder();
        if (readError.fileRowNum() > 0) {
            sb.append("fileRow=").append(readError.fileRowNum()).append(": ");
        }
        if (!readError.cellErrors().isEmpty()) {
            for (CellErrorResponse error : readError.cellErrors()) {
                sb.append("[column=").append(error.columnIndex());
                if (error.headerName() != null) {
                    sb.append(", header=").append(error.headerName());
                }
                sb.append(", value=").append(error.cellValue());
                sb.append(", message=").append(error.message()).append("] ");
            }
            return sb.toString().trim();
        }
        return sb.append(readError.messages()).toString();
    }

    private static String escapeHtml(Object value) {
        if (value == null) {
            return "";
        }
        return value.toString()
                .replace("&", "&amp;")
                .replace("<", "&lt;")
                .replace(">", "&gt;")
                .replace("\"", "&quot;");
    }
}
