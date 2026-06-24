package io.github.dornol.excelkit.example.app.book.adapter.out.persistence;

import org.apache.ibatis.annotations.Mapper;
import org.apache.ibatis.annotations.Select;

@Mapper
public interface BookMyBatisMapper {

    @Select("SELECT COUNT(*) FROM book")
    long countBooks();
}
