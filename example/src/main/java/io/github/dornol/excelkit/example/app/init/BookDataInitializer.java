package io.github.dornol.excelkit.example.app.init;

import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.boot.CommandLineRunner;
import org.springframework.jdbc.core.JdbcTemplate;
import org.springframework.stereotype.Component;

import java.util.ArrayList;
import java.util.List;
import java.util.Random;

@Component
public class BookDataInitializer implements CommandLineRunner {
    private static final Logger log = LoggerFactory.getLogger(BookDataInitializer.class);
    private final JdbcTemplate jdbcTemplate;
    private final Long dummyCount;
    private final Random random = new Random();

    public BookDataInitializer(JdbcTemplate jdbcTemplate, @Value("${example.dummy-count:1000000}") Long dummyCount) {
        this.jdbcTemplate = jdbcTemplate;
        this.dummyCount = dummyCount;
    }

    @Override
    public void run(String... args) {
        log.info("book data initializing...");
        int batchSize = 1000;
        long total = dummyCount;

        for (int i = 0; i < total; i += batchSize) {
            List<Object[]> batch = new ArrayList<>();
            for (int j = 0; j < batchSize; j++) {
                long id = i + j + 1L;
                batch.add(new Object[]{
                        randomText(1, 200), // title
                        randomText(1, 200), // subtitle
                        randomText(1, 200), // author
                        randomText(1, 200), // publisher
                        String.format("%013d", id), // isbn
                        randomText(5, 200), // title
                });
            }

            jdbcTemplate.batchUpdate("""
                INSERT INTO book (title, subtitle, author, publisher, isbn, description)
                VALUES (?, ?, ?, ?, ?, ?)
            """, batch);
        }

        log.warn("book data initialized: {}", total);
        log.warn("book data query: {}", jdbcTemplate.queryForObject("SELECT COUNT(*) FROM book", Long.class));
    }

    private String randomText(int minLength, int maxLength) {
        int len = minLength + random.nextInt(maxLength - minLength + 1);
        StringBuilder sb = new StringBuilder(len);
        for (int i = 0; i < len; i++) {
            char ch;
            int r = random.nextInt(36);
            if (r < 10) ch = (char) ('0' + r); // 숫자
            else ch = (char) ('a' + r - 10);   // 소문자
            sb.append(ch);
        }
        return sb.toString();
    }
}