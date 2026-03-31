package io.github.dornol.excelkit.example.app.showcase;

import io.github.dornol.excelkit.example.app.common.DownloadFileType;
import io.github.dornol.excelkit.example.app.common.DownloadUtil;
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
        return DownloadUtil.builder("type test excel", DownloadFileType.EXCEL)
                .body(handler::consumeOutputStream);
    }

}
