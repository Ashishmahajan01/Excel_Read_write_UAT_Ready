package com.excelDemo.excel.exception;


public class ExcelProcessingException extends RuntimeException {
    public ExcelProcessingException(String message) {
        super(message);
    }

    public ExcelProcessingException(String message, Throwable cause) {
        super(message, cause);
    }
}

