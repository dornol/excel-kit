package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.spring.ExcelKitResponse;
import io.github.dornol.excelkit.example.app.dto.TypeTestDto;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Controller;
import org.springframework.web.bind.annotation.GetMapping;
import org.springframework.web.servlet.mvc.method.annotation.StreamingResponseBody;

import java.util.stream.Stream;

@Controller
public class TypeTestController {

    @GetMapping("/download-excel-types")
    public ResponseEntity<StreamingResponseBody> downloadExcelTypes() {
        var handler = TypeTestExcelMapper.getHandler(Stream.generate(TypeTestDto::rand).limit(10000));
        return ExcelKitResponse.excel("type test excel")
                .body(handler::writeTo);
    }

}
