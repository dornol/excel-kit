package io.github.dornol.excelkit.example.app.repository;

import io.github.dornol.excelkit.example.app.dto.BookDto;
import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Options;
import org.apache.ibatis.annotations.Select;
import org.apache.ibatis.cursor.Cursor;

@Mapper
public interface BookMyBatisMapper {

    @Select("""
            SELECT id, title, subtitle, author, publisher, isbn, description
            FROM book
            WHERE id > 0
            ORDER BY id DESC
            """)
    @Options(fetchSize = 1000)
    Cursor<BookDto> getCursor();

}
