package com.aspose.cells_foss.core;

import java.time.LocalDateTime;
import java.time.Month;
import java.time.temporal.ChronoUnit;

/**
 * Converts between {@link LocalDateTime} values and Excel serial date numbers.
 */
public final class DateSerialConverter {

    private static final LocalDateTime WINDOWS_1900_EPOCH =
            LocalDateTime.of(1899, Month.DECEMBER, 31, 0, 0, 0);
    private static final LocalDateTime MAC_1904_EPOCH =
            LocalDateTime.of(1904, Month.JANUARY, 1, 0, 0, 0);
    // C# TimeSpan.TicksPerDay = 864_000_000_000 (100-ns ticks).
    // Java uses nanoseconds: 1 tick = 100 ns, so NANOS_PER_DAY = TicksPerDay * 100.
    private static final long NANOS_PER_DAY = 86_400_000_000_000L;

    /**
     * Initializes a new DateSerialConverter instance.
     */
    private DateSerialConverter() {}

    /**
     * Converts a {@link LocalDateTime} to an Excel serial date number.
     */
    public static double toSerial(LocalDateTime value, DateSystem dateSystem) {
        LocalDateTime baseDate = (dateSystem == DateSystem.MAC_1904) ? MAC_1904_EPOCH : WINDOWS_1900_EPOCH;
        long nanos = ChronoUnit.NANOS.between(baseDate, value);
        double serial = (double) nanos / NANOS_PER_DAY;
        // Handle the relevant branch before the state changes.
        if (dateSystem == DateSystem.WINDOWS_1900 && serial >= 60.0) {
            serial += 1.0;
        }
        return serial;
    }

    /**
     * Converts an Excel serial date number back to a {@link LocalDateTime}.
     */
    public static LocalDateTime fromSerial(double serial, DateSystem dateSystem) {
        // Handle the relevant branch before the state changes.
        if (dateSystem == DateSystem.WINDOWS_1900 && serial >= 60.0) {
            serial -= 1.0;
        }
        LocalDateTime baseDate = (dateSystem == DateSystem.MAC_1904) ? MAC_1904_EPOCH : WINDOWS_1900_EPOCH;
        long nanos = Math.round(serial * NANOS_PER_DAY);
        return baseDate.plus(nanos, ChronoUnit.NANOS);
    }
}