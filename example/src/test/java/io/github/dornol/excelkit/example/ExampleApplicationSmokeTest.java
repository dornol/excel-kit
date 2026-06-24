package io.github.dornol.excelkit.example;

import io.github.dornol.excelkit.example.app.book.adapter.out.persistence.BookMyBatisMapper;
import org.junit.jupiter.api.Test;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.boot.test.context.SpringBootTest;
import org.springframework.boot.webmvc.test.autoconfigure.AutoConfigureMockMvc;
import org.springframework.http.HttpHeaders;
import org.springframework.test.web.servlet.MockMvc;
import org.springframework.test.web.servlet.MvcResult;

import static org.junit.jupiter.api.Assertions.assertEquals;
import static org.junit.jupiter.api.Assertions.assertTrue;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.asyncDispatch;
import static org.springframework.test.web.servlet.request.MockMvcRequestBuilders.get;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.header;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.request;
import static org.springframework.test.web.servlet.result.MockMvcResultMatchers.status;

@AutoConfigureMockMvc
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

    @Autowired
    private MockMvc mockMvc;

    @Test
    void contextLoads() {
        assertEquals(0, bookMyBatisMapper.countBooks());
    }

    @Test
    void downloadExcel_returnsWorkbookResponse() throws Exception {
        MvcResult result = mockMvc.perform(get("/download-excel"))
                .andExpect(request().asyncStarted())
                .andReturn();

        MvcResult response = mockMvc.perform(asyncDispatch(result))
                .andExpect(status().isOk())
                .andExpect(header().string(HttpHeaders.CONTENT_TYPE,
                        "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"))
                .andExpect(header().string(HttpHeaders.CONTENT_DISPOSITION,
                        org.hamcrest.Matchers.containsString("book list excel.xlsx")))
                .andReturn();

        byte[] body = response.getResponse().getContentAsByteArray();
        assertTrue(body.length > 4);
        assertEquals('P', body[0]);
        assertEquals('K', body[1]);
    }
}
