package com.aspose.cells_foss;

import static org.junit.jupiter.api.Assertions.*;

/**
 * Assertion helpers mirroring the C# AssertEx class from the shared test infrastructure.
 */
public final class AssertEx {

    /**
     * Verifies that assert ex.
     */
    private AssertEx() {}

    /**
     * Verifies that assert true.
     * @param condition condition
     */
    public static void assertTrue(boolean condition) {
        assertTrue(condition, null);
    }

    /**
     * Verifies that assert true.
     * @param condition condition
     * @param message message
     */
    public static void assertTrue(boolean condition, String message) {
        org.junit.jupiter.api.Assertions.assertTrue(condition, message != null ? message : "Expected condition to be true.");
    }

    /**
     * Verifies that assert false.
     * @param condition condition
     */
    public static void assertFalse(boolean condition) {
        assertFalse(condition, null);
    }

    /**
     * Verifies that assert false.
     * @param condition condition
     * @param message message
     */
    public static void assertFalse(boolean condition, String message) {
        org.junit.jupiter.api.Assertions.assertFalse(condition, message != null ? message : "Expected condition to be false.");
    }

    /**
     * Verifies that assert equal.
     * @param expected expected
     * @param actual actual
     */
    public static <T> void assertEqual(T expected, T actual) {
        assertEqual(expected, actual, null);
    }

    /**
     * Verifies that assert equal.
     * @param expected expected
     * @param actual actual
     * @param message message
     */
    public static <T> void assertEqual(T expected, T actual, String message) {
        assertEquals(expected, actual, message != null ? message : "Expected '" + expected + "', got '" + actual + "'.");
    }

    /**
     * Verifies that assert not null.
     * @param value value to apply
     */
    public static void assertNotNull(Object value) {
        assertNotNull(value, null);
    }

    /**
     * Verifies that assert not null.
     * @param value value to apply
     * @param message message
     */
    public static void assertNotNull(Object value, String message) {
        org.junit.jupiter.api.Assertions.assertNotNull(value, message != null ? message : "Expected value to be non-null.");
    }

    /**
     * Verifies that assert null.
     * @param value value to apply
     */
    public static void assertNull(Object value) {
        assertNull(value, null);
    }

    /**
     * Verifies that assert null.
     * @param value value to apply
     * @param message message
     */
    public static void assertNull(Object value, String message) {
        org.junit.jupiter.api.Assertions.assertNull(value, message != null ? message : "Expected value to be null, got '" + value + "'.");
    }

    /**
     * Verifies that assert contains.
     * @param expectedSubstring expected substring
     * @param actual actual
     */
    public static void assertContains(String expectedSubstring, String actual) {
        assertContains(expectedSubstring, actual, null);
    }

    /**
     * Verifies that assert contains.
     * @param expectedSubstring expected substring
     * @param actual actual
     * @param message message
     */
    public static void assertContains(String expectedSubstring, String actual, String message) {
        org.junit.jupiter.api.Assertions.assertTrue(
                actual != null && actual.contains(expectedSubstring),
                message != null ? message : "Expected '" + actual + "' to contain '" + expectedSubstring + "'.");
    }

    /**
     * Verifies that assert throws.
     * @param expectedType expected type
     * @param action action
     */
    public static <T extends Throwable> T assertThrows(Class<T> expectedType, Executable action) {
        return assertThrows(expectedType, action, null);
    }

    /**
     * Verifies that assert throws.
     * @param expectedType expected type
     * @param action action
     * @param message message
     */
    public static <T extends Throwable> T assertThrows(Class<T> expectedType, Executable action, String message) {
        return org.junit.jupiter.api.Assertions.assertThrows(expectedType, action::execute,
                message != null ? message : "Expected " + expectedType.getSimpleName() + " to be thrown.");
    }

    /**
     * Test coverage for Executable behaviors.
     */
    @FunctionalInterface
    public interface Executable {
        void execute() throws Exception;
    }
}
