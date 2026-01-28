package com.excelDemo.excel.service;

import com.excelDemo.excel.dto.FileResponse;
import com.excelDemo.excel.exception.ExcelProcessingException;
import com.excelDemo.excel.exception.InvalidFileFormatException;
import com.excelDemo.excel.model.UserData;
import com.excelDemo.excel.repository.UserDataRepository;
import com.excelDemo.excel.util.ExcelIUtil;
import lombok.RequiredArgsConstructor;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.time.LocalDateTime;
import java.util.*;

@Service
@RequiredArgsConstructor
public class ExcelUploadService {

    private final UserDataRepository userDataRepository;
    private static final String USER_NAME = "username";
    public static final String EMAIL = "email";
    public static final String AGE = "age";
    public static final String DEPARTMENT = "department";
    public static final String SALARY = "salary";
    public static final String IS_ACTIVATE = "is_active";
    private static final List<String> EXPECTED_HEADERS = Arrays.asList(
            USER_NAME, EMAIL, AGE, DEPARTMENT, SALARY, IS_ACTIVATE
    );
    private static final Map<String, Integer> HEADER_MAP = new HashMap<>();

    public FileResponse processExcelFile(MultipartFile file) throws IOException {
        List<String> errors = new ArrayList<>();
        int recordsProcessed = 0;

        if (file == null || file.isEmpty()) {
            errors.add("File is empty or null");
            return createErrorResponse(file, errors, 0);
        }

        List<UserData> userDataList = new ArrayList<>();

        try (XSSFWorkbook workbook = new XSSFWorkbook(file.getInputStream())) {
            Sheet sheet = workbook.getSheetAt(0);

            if (sheet == null) {
                errors.add("Excel file does not contain any sheets");
                return createErrorResponse(file, errors, 0);
            }

            Iterator<Row> rows = sheet.iterator();

            // Process header row
            if (!rows.hasNext()) {
                errors.add("Excel file is empty");
                return createErrorResponse(file, errors, 0);
            }

            Row headerRow = rows.next();
            List<String> headerValues = extractHeaderValues(headerRow);

            // Perform header validation
            try {
                ExcelIUtil.headerChecks(headerValues, HEADER_MAP, EXPECTED_HEADERS);
            } catch (Exception e) {
                errors.add("Invalid header format: " + e.getMessage());
                return createErrorResponse(file, errors, 0);
            }

            // Get total row count (excluding header) to check limit (prevent DoS)
            int lastRowNum = sheet.getLastRowNum();
            int totalDataRows = lastRowNum; // getLastRowNum() returns 0-based index, excluding header
            
            // Validate row limit to prevent DoS attacks
            if (totalDataRows > ExcelIUtil.MAX_ROWS_LIMIT) {
                errors.add("File contains too many rows (" + totalDataRows + "). Maximum allowed: " + ExcelIUtil.MAX_ROWS_LIMIT);
                return createErrorResponse(file, errors, 0);
            }
            
            // Process data rows
            int rowNumber = 1; // Start counting after header
            while (rows.hasNext()) {
                Row currentRow = rows.next();
                rowNumber++;

                try {
                    UserData userData = createUserDataFromRow(currentRow);
                    userDataList.add(userData);
                } catch (Exception e) {
                    errors.add("Error processing row " + rowNumber + ": " + e.getMessage());
                }
            }

            if (userDataList.isEmpty()) {
                errors.add("No valid data found in the Excel file");
                return createErrorResponse(file, errors, 0);
            }

            // Save to database
            try {
                List<UserData> savedUsers = userDataRepository.saveAll(userDataList);
                recordsProcessed = savedUsers.size();

                if (!errors.isEmpty()) {
                    return new FileResponse(
                            file.getOriginalFilename(),
                            (int) file.getSize(),
                            LocalDateTime.now(),
                            "Partial success - " + recordsProcessed + " records processed with " + errors.size() + " errors",
                            errors,
                            recordsProcessed
                    );
                } else {
                    return new FileResponse(
                            file.getOriginalFilename(),
                            (int) file.getSize(),
                            LocalDateTime.now(),
                            "Success - " + recordsProcessed + " records processed successfully",
                            null,
                            recordsProcessed
                    );
                }

            } catch (Exception e) {
                errors.add("Error saving data to database: " + e.getMessage());
                return createErrorResponse(file, errors, 0);
            }

        } catch (IOException e) {
            errors.add("Error reading Excel file: " + e.getMessage());
            return createErrorResponse(file, errors, 0);
        } catch (Exception e) {
            errors.add("Error processing Excel file: " + e.getMessage());
            return createErrorResponse(file, errors, 0);
        }
    }

    private FileResponse createErrorResponse(MultipartFile file, List<String> errors, int recordsProcessed) {
        return new FileResponse(
                file != null ? file.getOriginalFilename() : null,
                file != null ? (int) file.getSize() : 0,
                LocalDateTime.now(),
                "Failed",
                errors,
                recordsProcessed
        );
    }

    private UserData createUserDataFromRow(Row row) {
        UserData userData = new UserData();

        try {
            // Map Excel columns to entity fields with sanitization and validation
            String username = safeString(getStringCellValue(row.getCell(0)));
            ExcelIUtil.validateUsername(username);
            userData.setUsername(username);
            
            String email = safeString(getStringCellValue(row.getCell(1)));
            ExcelIUtil.validateEmail(email);
            userData.setEmail(email);
            
            Integer age = getIntegerCellValue(row.getCell(2));
            ExcelIUtil.validateAge(age);
            userData.setAge(age);
            
            String department = safeString(getStringCellValue(row.getCell(3)));
            ExcelIUtil.validateDepartment(department);
            userData.setDepartment(department);
            
            Double salary = getDoubleCellValue(row.getCell(4));
            ExcelIUtil.validateSalary(salary);
            userData.setSalary(salary);
            
            userData.setIsActive(getBooleanCellValue(row.getCell(5)));

        } catch (com.excelDemo.excel.exception.ValidationException e) {
            // Re-throw validation exceptions as-is (they return 400)
            throw e;
        } catch (Exception e) {
            throw new ExcelProcessingException("Error mapping row data to entity: " + e.getMessage(), e);
        }

        return userData;
    }

    // Helper methods remain the same...
    private String getStringCellValue(Cell cell) {
        if (cell == null) return null;

        return switch (cell.getCellType()) {
            case STRING -> cell.getStringCellValue();
            case NUMERIC -> String.valueOf(cell.getNumericCellValue());
            case BOOLEAN -> String.valueOf(cell.getBooleanCellValue());
            default -> null;
        };
    }

    private Integer getIntegerCellValue(Cell cell) {
        if (cell == null) return null;

        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return (int) cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                return Integer.parseInt(cell.getStringCellValue());
            }
        } catch (Exception e) {
            throw new ExcelProcessingException("Invalid integer value in cell");
        }
        return null;
    }

    private Double getDoubleCellValue(Cell cell) {
        if (cell == null) return null;

        try {
            if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                return Double.parseDouble(cell.getStringCellValue());
            }
        } catch (Exception e) {
            throw new ExcelProcessingException("Invalid decimal value in cell");
        }
        return null;
    }

    private Boolean getBooleanCellValue(Cell cell) {
        if (cell == null) return null;

        try {
            if (cell.getCellType() == CellType.BOOLEAN) {
                return cell.getBooleanCellValue();
            } else if (cell.getCellType() == CellType.STRING) {
                return Boolean.parseBoolean(cell.getStringCellValue());
            } else if (cell.getCellType() == CellType.NUMERIC) {
                return cell.getNumericCellValue() == 1;
            }
        } catch (Exception e) {
            throw new ExcelProcessingException("Invalid boolean value in cell");
        }
        return null;
    }

    private LocalDateTime getDateTimeCellValue(Cell cell) {
        if (cell == null) return null;

        try {
            if (cell.getCellType() == CellType.NUMERIC && DateUtil.isCellDateFormatted(cell)) {
                return cell.getLocalDateTimeCellValue();
            }
        } catch (Exception e) {
            throw new ExcelProcessingException("Invalid date/time value in cell");
        }
        return null;
    }
    private List<String> extractHeaderValues(Row headerRow) {
        List<String> headerValues = new ArrayList<>();
        Iterator<Cell> cellIterator = headerRow.cellIterator();

        while (cellIterator.hasNext()) {
            Cell cell = cellIterator.next();
            String cellValue = getStringCellValue(cell);
            headerValues.add(cellValue != null ? cellValue.trim().toLowerCase() : "");
        }
        return headerValues;
    }

    public static String safeString(String value) {
        return value != null ? ExcelIUtil.sanitize(value).trim() : null;
    }
}
