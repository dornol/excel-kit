package io.github.dornol.excelkit.example.app.dto;

import java.math.BigDecimal;
import java.math.RoundingMode;
import java.security.SecureRandom;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.LocalTime;
import java.time.ZoneOffset;
import java.util.UUID;

public record TypeTestDto(
        String aString,
        Long aLong,
        Integer anInteger,
        LocalDateTime aLocalDateTime,
        LocalDate aLocalDate,
        LocalTime aLocalTime,
        Double aDouble,
        Float aFloat,
        Boolean aBoolean,
        BigDecimal aLongBigDecimal,
        BigDecimal aDoubleBigDecimal
) {

    public static TypeTestDto rand() {
        SecureRandom random = new SecureRandom();
        long minEpochSecond = LocalDateTime.of(2000, 1, 1, 0, 0).toEpochSecond(ZoneOffset.UTC);
        long maxEpochSecond = LocalDateTime.of(2030, 12, 31, 23, 59).toEpochSecond(ZoneOffset.UTC);

        long randomEpoch = minEpochSecond + random.nextLong(maxEpochSecond - minEpochSecond);
        LocalDateTime randomDateTime = LocalDateTime.ofEpochSecond(randomEpoch, 0, ZoneOffset.UTC);
        return new TypeTestDto(
                UUID.randomUUID().toString(),
                random.nextLong(),
                random.nextInt(),
                randomDateTime,
                randomDateTime.toLocalDate(),
                randomDateTime.toLocalTime(),
                random.nextDouble() * 50000,
                random.nextFloat() * 50000,
                random.nextBoolean(),
                BigDecimal.valueOf(random.nextLong()),
                BigDecimal.valueOf(random.nextDouble() * 100000).setScale(2, RoundingMode.HALF_UP));
    }

}
