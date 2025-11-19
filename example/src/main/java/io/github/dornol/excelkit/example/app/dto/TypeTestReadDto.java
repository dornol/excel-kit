package io.github.dornol.excelkit.example.app.dto;

import java.math.BigDecimal;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;

public class TypeTestReadDto {
    private Long no;
    private String stringVal;
    private Long longVal;
    private Integer integer;
    private LocalDateTime localDateTime;
    private LocalDate localDate;
    private LocalTime localTime;
    private Double doubleVal;
    private Float floatVal;
    private Boolean booleanVal;
    private BigDecimal longBigDecimal;
    private BigDecimal doubleBigDecimal;

    public Long getNo() {
        return no;
    }

    public void setNo(Long no) {
        this.no = no;
    }

    public String getStringVal() {
        return stringVal;
    }

    public void setStringVal(String stringVal) {
        this.stringVal = stringVal;
    }

    public Long getLongVal() {
        return longVal;
    }

    public void setLongVal(Long longVal) {
        this.longVal = longVal;
    }

    public Integer getInteger() {
        return integer;
    }

    public void setInteger(Integer integer) {
        this.integer = integer;
    }

    public LocalDateTime getLocalDateTime() {
        return localDateTime;
    }

    public void setLocalDateTime(LocalDateTime localDateTime) {
        this.localDateTime = localDateTime;
    }

    public LocalDate getLocalDate() {
        return localDate;
    }

    public void setLocalDate(LocalDate localDate) {
        this.localDate = localDate;
    }

    public LocalTime getLocalTime() {
        return localTime;
    }

    public void setLocalTime(LocalTime localTime) {
        this.localTime = localTime;
    }

    public Double getDoubleVal() {
        return doubleVal;
    }

    public void setDoubleVal(Double doubleVal) {
        this.doubleVal = doubleVal;
    }

    public Float getFloatVal() {
        return floatVal;
    }

    public void setFloatVal(Float floatVal) {
        this.floatVal = floatVal;
    }

    public Boolean getBooleanVal() {
        return booleanVal;
    }

    public void setBooleanVal(Boolean booleanVal) {
        this.booleanVal = booleanVal;
    }

    public BigDecimal getLongBigDecimal() {
        return longBigDecimal;
    }

    public void setLongBigDecimal(BigDecimal longBigDecimal) {
        this.longBigDecimal = longBigDecimal;
    }

    public BigDecimal getDoubleBigDecimal() {
        return doubleBigDecimal;
    }

    public void setDoubleBigDecimal(BigDecimal doubleBigDecimal) {
        this.doubleBigDecimal = doubleBigDecimal;
    }

    @Override
    public String toString() {
        return "TypeTestReadDto{" +
                "no=" + no +
                ", aString='" + stringVal + '\'' +
                ", aLong=" + longVal +
                ", anInteger=" + integer +
                ", aLocalDateTime=" + localDateTime +
                ", aLocalDate=" + localDate +
                ", aLocalTime=" + localTime +
                ", aDouble=" + doubleVal +
                ", aFloat=" + floatVal +
                ", aBoolean=" + booleanVal +
                ", aLongBigDecimal=" + longBigDecimal +
                ", aDoubleBigDecimal=" + doubleBigDecimal +
                '}';
    }
}
