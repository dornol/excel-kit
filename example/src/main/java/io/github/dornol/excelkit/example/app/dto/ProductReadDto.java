package io.github.dornol.excelkit.example.app.dto;

public class ProductReadDto {
    private String name;
    private String category;
    private Integer price;
    private Integer quantity;
    private Double discount;
    private String url;

    public String getName() { return name; }
    public void setName(String name) { this.name = name; }

    public String getCategory() { return category; }
    public void setCategory(String category) { this.category = category; }

    public Integer getPrice() { return price; }
    public void setPrice(Integer price) { this.price = price; }

    public Integer getQuantity() { return quantity; }
    public void setQuantity(Integer quantity) { this.quantity = quantity; }

    public Double getDiscount() { return discount; }
    public void setDiscount(Double discount) { this.discount = discount; }

    public String getUrl() { return url; }
    public void setUrl(String url) { this.url = url; }

    @Override
    public String toString() {
        return "ProductReadDto{name='%s', category='%s', price=%d, quantity=%d, discount=%s, url='%s'}"
                .formatted(name, category, price, quantity, discount, url);
    }
}
