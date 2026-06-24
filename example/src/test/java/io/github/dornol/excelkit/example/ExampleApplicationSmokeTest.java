package io.github.dornol.excelkit.example;

import io.github.dornol.excelkit.example.app.book.adapter.out.persistence.BookMyBatisMapper;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;

import static org.junit.jupiter.api.Assertions.assertEquals;

@SpringBootTest(properties = {
        "spring.docker.compose.enabled=false",
        "spring.datasource.url=jdbc:h2:mem:example-smoke;MODE=MariaDB;DB_CLOSE_DELAY=-1;DATABASE_TO_LOWER=TRUE",
        "spring.datasource.driver-class-name=org.h2.Driver",
        "spring.datasource.username=sa",
        "spring.datasource.password=",
        "spring.jpa.hibernate.ddl-auto=create-drop",
        "spring.jpa.show-sql=false",
        "example.dummy-count=0"
})
class ExampleApplicationSmokeTest {

    @Autowired
    private BookMyBatisMapper bookMyBatisMapper;

    @Test
    void contextLoads() {
        assertEquals(0, bookMyBatisMapper.countBooks());
    }
}
