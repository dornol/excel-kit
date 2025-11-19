package io.github.dornol.excelkit.example.app.excel;

import io.github.dornol.excelkit.example.app.dto.TypeTestDto;
import io.github.dornol.excelkit.example.app.dto.TypeTestReadDto;
import io.github.dornol.excelkit.excel.*;
import jakarta.validation.Validator;

import java.io.InputStream;
import java.security.SecureRandom;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.util.UUID;
import java.util.stream.Stream;

public class TypeTestExcelMapper {
    private TypeTestExcelMapper() {
        /* empty */
    }

    public static ExcelHandler getHandler(Stream<TypeTestDto> stream) {
        var date = LocalDate.now();
        SecureRandom random = new SecureRandom();
        return new ExcelWriter<TypeTestDto>(147, 252, 42, 1000000)
                .column("no", (rowData, cursor) -> cursor.getCurrentTotal()).type(ExcelDataType.INTEGER)
                .column("string", TypeTestDto::aString)
                .column("long", TypeTestDto::aLong).type(ExcelDataType.LONG)
                .column("integer", TypeTestDto::anInteger).type(ExcelDataType.INTEGER)
                .column("localdatetime", TypeTestDto::aLocalDateTime).type(ExcelDataType.DATETIME)
                .column("localdate", TypeTestDto::aLocalDate).type(ExcelDataType.DATE)
                .column("localtime", r -> LocalDateTime.of(date, r.aLocalTime())).type(ExcelDataType.TIME)
                .column("double", TypeTestDto::aDouble).type(ExcelDataType.DOUBLE)
                .column("double_percent", TypeTestDto::aDouble).type(ExcelDataType.DOUBLE_PERCENT)
                .column("float", TypeTestDto::aFloat).type(ExcelDataType.FLOAT)
                .column("float_percent", TypeTestDto::aFloat).type(ExcelDataType.FLOAT_PERCENT)
                .column("boolean", TypeTestDto::aBoolean).type(ExcelDataType.BOOLEAN_TO_YN)
                .column("bigdecimal_long", TypeTestDto::aLongBigDecimal).type(ExcelDataType.BIG_DECIMAL_TO_LONG)
                .column("bigdecimal_double", TypeTestDto::aDoubleBigDecimal).type(ExcelDataType.BIG_DECIMAL_TO_DOUBLE)
                .column("currency", TypeTestDto::aDoubleBigDecimal).type(ExcelDataType.BIG_DECIMAL_TO_DOUBLE).format(ExcelDataFormat.CURRENCY_USD.getFormat())
                .constColumn("const", random.nextLong()).type(ExcelDataType.LONG)
                .constColumn("const_string", UUID.randomUUID().toString())
                .write(stream);
    }

    public static ExcelReadHandler<TypeTestReadDto> getReadHandler(InputStream inputStream, Validator validator) {
        return new ExcelReader<>(TypeTestReadDto::new, validator)
                .column((d, c) -> d.setNo(c.asLong()))
                .column((d, c) -> d.setStringVal(c.asString()))
                .column((d, c) -> d.setLongVal(c.asLong()))
                .column((d, c) -> d.setInteger(c.asInt()))
                .column((d, c) -> d.setLocalDateTime(c.asLocalDateTime()))
                .column((d, c) -> d.setLocalDate(c.asLocalDate()))
                .column((d, c) -> d.setLocalTime(c.asLocalTime()))
                .column((d, c) -> d.setDoubleVal(c.asDouble()))
                .column((d, c) -> d.setDoubleVal(c.asDouble()))
                .column((d, c) -> d.setFloatVal(c.asFloat()))
                .column((d, c) -> d.setFloatVal(c.asFloat()))
                .column((d, c) -> d.setBooleanVal(c.asBoolean()))
                .column((d, c) -> d.setLongBigDecimal(c.asBigDecimal()))
                .column((d, c) -> d.setDoubleBigDecimal(c.asBigDecimal()))
                .column((d, c) -> d.setDoubleBigDecimal(c.asBigDecimal()))
                .build(inputStream);
    }
}
