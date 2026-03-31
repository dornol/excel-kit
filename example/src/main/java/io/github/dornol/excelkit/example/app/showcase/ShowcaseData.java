package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.dto.ProductDto;
import io.github.dornol.excelkit.example.app.dto.ProductReadDto;
import io.github.dornol.excelkit.excel.ExcelDataType;
import io.github.dornol.excelkit.shared.ExcelKitSchema;

import java.util.List;
import java.util.stream.Stream;

/**
 * Shared data helpers and schema definitions for Showcase controllers.
 */
public final class ShowcaseData {

    public static final ExcelKitSchema<ProductReadDto> PRODUCT_SCHEMA = ExcelKitSchema.<ProductReadDto>builder()
            .column("Name", ProductReadDto::getName, (p, cell) -> p.setName(cell.asString()))
            .column("Category", ProductReadDto::getCategory, (p, cell) -> p.setCategory(cell.asString()))
            .column("Price", ProductReadDto::getPrice, (p, cell) -> p.setPrice(cell.asInt()),
                    c -> c.type(ExcelDataType.INTEGER).format("#,##0"))
            .column("Quantity", ProductReadDto::getQuantity, (p, cell) -> p.setQuantity(cell.asInt()),
                    c -> c.type(ExcelDataType.INTEGER))
            .column("Discount", ProductReadDto::getDiscount, (p, cell) -> p.setDiscount(cell.asDouble()),
                    c -> c.type(ExcelDataType.DOUBLE_PERCENT))
            .build();

    public static List<ProductDto> sampleProducts() {
        return Stream.generate(ProductDto::random).limit(20).toList();
    }

    public static ProductReadDto toReadDto(ProductDto p) {
        var dto = new ProductReadDto();
        dto.setName(p.name());
        dto.setCategory(p.category());
        dto.setPrice(p.price());
        dto.setQuantity(p.quantity());
        dto.setDiscount(p.discount());
        dto.setUrl(p.url());
        return dto;
    }

    private ShowcaseData() {
    }
}
