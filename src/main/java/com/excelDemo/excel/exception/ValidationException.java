package com.excelDemo.excel.exception;

/**
 * Exception thrown when input validation fails (e.g., sanitization detects malicious content).
 * Returns HTTP 400 Bad Request instead of 500 Internal Server Error.
 */
public class ValidationException extends RuntimeException {
    public ValidationException(String message) {
        super(message);
    }

    public ValidationException(String message, Throwable cause) {
        super(message, cause);
    }
}
