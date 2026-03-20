package io.github.dornol.excelkit.example.app.service;

import io.github.dornol.excelkit.example.app.dto.BookDto;
import io.github.dornol.excelkit.example.app.excel.BookExcelMapper;
import io.github.dornol.excelkit.example.app.repository.BookMyBatisMapper;
import io.github.dornol.excelkit.excel.ExcelHandler;
import org.apache.ibatis.cursor.Cursor;
import org.springframework.stereotype.Service;
import org.springframework.transaction.annotation.Transactional;

import java.util.Spliterator;
import java.util.Spliterators;
import java.util.stream.Stream;
import java.util.stream.StreamSupport;

@Service
public class BookMyBatisService {

    private final BookMyBatisMapper bookMyBatisMapper;

    public BookMyBatisService(BookMyBatisMapper bookMyBatisMapper) {
        this.bookMyBatisMapper = bookMyBatisMapper;
    }

    @Transactional(readOnly = true)
    public ExcelHandler getExcelHandler() {
        Cursor<BookDto> cursor = bookMyBatisMapper.getCursor();
        Stream<BookDto> stream = StreamSupport.stream(
                Spliterators.spliteratorUnknownSize(cursor.iterator(), Spliterator.ORDERED),
                false
        ).onClose(() -> {
            try {
                cursor.close();
            } catch (Exception e) {
                // ignore
            }
        });
        return BookExcelMapper.getHandler(stream);
    }

}
