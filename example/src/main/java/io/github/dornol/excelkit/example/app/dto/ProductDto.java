package io.github.dornol.excelkit.example.app.dto;

import java.security.SecureRandom;

public record ProductDto(
        String name,
        String category,
        int price,
        int quantity,
        double discount,
        String url
) {

    private static final SecureRandom RANDOM = new SecureRandom();
    private static final String[] NAMES = {
            "Wireless Mouse", "Mechanical Keyboard", "USB-C Hub", "Monitor Stand",
            "Webcam HD", "Bluetooth Speaker", "Laptop Bag", "Screen Protector",
            "Phone Charger", "HDMI Cable", "Mouse Pad", "Desk Lamp"
    };
    private static final String[] CATEGORIES = {"Electronics", "Accessories", "Office", "Peripherals"};

    public static ProductDto random() {
        String name = NAMES[RANDOM.nextInt(NAMES.length)];
        String category = CATEGORIES[RANDOM.nextInt(CATEGORIES.length)];
        int price = (RANDOM.nextInt(50) + 1) * 1000;
        int quantity = RANDOM.nextInt(100) + 1;
        double discount = RANDOM.nextInt(30) / 100.0;
        String url = "https://example.com/product/" + name.toLowerCase().replace(" ", "-");
        return new ProductDto(name, category, price, quantity, discount, url);
    }
}
