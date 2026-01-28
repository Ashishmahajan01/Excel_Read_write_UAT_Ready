package com.excelDemo.excel.util;


import com.excelDemo.excel.exception.ValidationException;
import org.apache.coyote.BadRequestException;
import org.apache.poi.ooxml.POIXMLProperties;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;
import org.springframework.http.ContentDisposition;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.web.multipart.MultipartFile;
import org.springframework.web.util.HtmlUtils;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.Map;
import java.util.regex.Matcher;
import java.util.regex.Pattern;

public class ExcelIUtil {
    private static final String TYPE = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
    // XLSX file signature: PK (ZIP format) - first 4 bytes should be 50 4B 03 04 or 50 4B 05 06 or 50 4B 07 08
    private static final byte[] XLSX_SIGNATURE_1 = {(byte) 0x50, (byte) 0x4B, 0x03, 0x04};
    private static final byte[] XLSX_SIGNATURE_2 = {(byte) 0x50, (byte) 0x4B, 0x05, 0x06};
    private static final byte[] XLSX_SIGNATURE_3 = {(byte) 0x50, (byte) 0x4B, 0x07, 0x08};
    private static final Pattern FORMULA_PATTERN = Pattern.compile("^[=+\\-@]");
    private static final Pattern HTML_PATTERN = Pattern.compile("^[<>\"'&]");
    private static final String CACHE_HEADER = "no-store,no-cache,must-revalidate";
    private static final String CACHE = "no-cache";
    private static final String CONTEXT_TYPE = "X-Content-Type-Options"; // Fixed casing for pen-test compliance
    private static final String SNIFF = "nosniff";
    
    // Field validation constants
    public static final int MAX_USERNAME_LENGTH = 100;
    public static final int MAX_EMAIL_LENGTH = 255;
    public static final int MAX_DEPARTMENT_LENGTH = 100;
    public static final int MIN_AGE = 0;
    public static final int MAX_AGE = 150;
    public static final double MIN_SALARY = 0.0;
    public static final double MAX_SALARY = 10_000_000.0; // 10 million max
    public static final int MAX_ROWS_LIMIT = 10000; // Prevent DoS attacks with huge files
    private static final Pattern EMAIL_PATTERN = Pattern.compile(
        "^[A-Za-z0-9+_.-]+@[A-Za-z0-9.-]+\\.[A-Za-z]{2,}$"
    );
    private static final String SCRIPT_TAG_PATTERN = "(?i)<\\s*script\\b[^>]*>(.*?)<\\s*/\\s*script\\s*>";
    private static final String INLINE_EVENT_HANDLER_PATTERN = "(?i)on\\w+\\s*=\\s*['\"]?[^'\"]+['\"]?";
    private static final String JS_URI_PATTERN = "(?i)(javascript:|data:text/html)";
    private static final String DANGEROUS_TAGS_PATTERN = "(?i)<\\s*(iframe|embed|svg|object)\\b[^>]*>";
    private static final String IMG_EVENT_PATTERN = "(?i)<\\s*img\\b[^>]*on\\w+\\s*=";
    public static final String CSV_SUFFIX = ".csv";
    private static final String DANGEROUS_JS_APIS_PATTERN = "(?i)(document\\.cookie|document\\.write|window\\.location|eval\\s*\\(|String\\.fromCharCode\\s*\\()";
    static final Logger logger = LoggerFactory.getLogger(ExcelIUtil.class);

    private static final List<Pattern> XSS_PATTERNS = Arrays.asList(
            Pattern.compile(SCRIPT_TAG_PATTERN),
            Pattern.compile(INLINE_EVENT_HANDLER_PATTERN),
            Pattern.compile(JS_URI_PATTERN),
            Pattern.compile(DANGEROUS_TAGS_PATTERN),
            Pattern.compile(IMG_EVENT_PATTERN),
            Pattern.compile(DANGEROUS_JS_APIS_PATTERN)
    );
    private static final CharSequence COMMA = ", ";

    /**
     * Validates input for script injection patterns.
     * Throws ValidationException (400) instead of IllegalArgumentException (500) for better pen-test compliance.
     */
    public static void isScriptInjection(String input) {
        if (input == null) return;
        
        for (Pattern pattern : XSS_PATTERNS) {
            Matcher matcher = pattern.matcher(input);
            if (matcher.find()) {
                // Return 400 Bad Request instead of 500 for validation errors
                throw new ValidationException("Malicious content detected. Input contains potentially dangerous script patterns.");
            }
        }
    }

    /**
     * Validates file format by checking Content-Type and file signature (magic bytes).
     * This prevents attackers from uploading malicious files with fake Content-Type headers.
     */
    public static boolean hasExcelFormat(MultipartFile file) {
        if (file == null || file.isEmpty()) {
            return false;
        }
        
        // Check Content-Type
        String contentType = file.getContentType();
        if (!TYPE.equals(contentType)) {
            return false;
        }
        
        // Verify file signature (magic bytes) to ensure it's actually an XLSX file
        try {
            byte[] fileBytes = file.getBytes();
            if (fileBytes.length < 4) {
                return false;
            }
            
            // Check for XLSX signature (ZIP format)
            boolean matchesSignature1 = Arrays.equals(
                Arrays.copyOfRange(fileBytes, 0, 4), XLSX_SIGNATURE_1
            );
            boolean matchesSignature2 = Arrays.equals(
                Arrays.copyOfRange(fileBytes, 0, 4), XLSX_SIGNATURE_2
            );
            boolean matchesSignature3 = Arrays.equals(
                Arrays.copyOfRange(fileBytes, 0, 4), XLSX_SIGNATURE_3
            );
            
            return matchesSignature1 || matchesSignature2 || matchesSignature3;
        } catch (IOException e) {
            logger.error("Error reading file bytes for signature validation", e);
            return false;
        }
    }
    
    /**
     * Validates file signature by reading first bytes from InputStream.
     * More memory-efficient for large files.
     */
    public static boolean validateExcelSignature(MultipartFile file) {
        try {
            byte[] header = new byte[4];
            int bytesRead = file.getInputStream().read(header);
            if (bytesRead < 4) {
                return false;
            }
            
            boolean matchesSignature1 = Arrays.equals(header, XLSX_SIGNATURE_1);
            boolean matchesSignature2 = Arrays.equals(header, XLSX_SIGNATURE_2);
            boolean matchesSignature3 = Arrays.equals(header, XLSX_SIGNATURE_3);
            
            return matchesSignature1 || matchesSignature2 || matchesSignature3;
        } catch (IOException e) {
            logger.error("Error validating Excel signature", e);
            return false;
        }
    }

    /**
     * Sanitizes input string to prevent XSS, script injection, and formula injection.
     * This is the core sanitization method used throughout the application.
     * 
     * @param input Raw input string from Excel cell
     * @return Sanitized string safe for storage and display
     * @throws ValidationException if malicious content is detected
     */
    public static String sanitize(String input) {
        if (input == null) return "";

        // Unescape HTML entities first (e.g., &lt;script&gt; becomes <script>)
        // This allows us to detect encoded malicious content
        String sanitized = HtmlUtils.htmlUnescape(input);
        
        // Check for script injection patterns (throws ValidationException if found)
        isScriptInjection(sanitized);

        // Prevent Excel formula injection by prefixing dangerous characters
        // Formulas starting with =, +, -, @ can execute code
        // HTML-like patterns (<, >, ", ', &) can be used for XSS
        if (FORMULA_PATTERN.matcher(sanitized).find() || HTML_PATTERN.matcher(sanitized).find()) {
            sanitized = "'" + sanitized; // Prefix with single quote to make Excel treat as text
        }

        return sanitized.trim();
    }
    
    /**
     * Validates email format using regex pattern.
     */
    public static void validateEmail(String email) {
        if (email == null || email.trim().isEmpty()) {
            throw new ValidationException("Email cannot be empty");
        }
        
        if (email.length() > MAX_EMAIL_LENGTH) {
            throw new ValidationException("Email exceeds maximum length of " + MAX_EMAIL_LENGTH + " characters");
        }
        
        if (!EMAIL_PATTERN.matcher(email).matches()) {
            throw new ValidationException("Invalid email format: " + email);
        }
    }
    
    /**
     * Validates username field.
     */
    public static void validateUsername(String username) {
        if (username == null || username.trim().isEmpty()) {
            throw new ValidationException("Username cannot be empty");
        }
        
        if (username.length() > MAX_USERNAME_LENGTH) {
            throw new ValidationException("Username exceeds maximum length of " + MAX_USERNAME_LENGTH + " characters");
        }
    }
    
    /**
     * Validates age field.
     */
    public static void validateAge(Integer age) {
        if (age == null) {
            throw new ValidationException("Age cannot be null");
        }
        
        if (age < MIN_AGE || age > MAX_AGE) {
            throw new ValidationException("Age must be between " + MIN_AGE + " and " + MAX_AGE);
        }
    }
    
    /**
     * Validates salary field.
     */
    public static void validateSalary(Double salary) {
        if (salary == null) {
            throw new ValidationException("Salary cannot be null");
        }
        
        if (salary < MIN_SALARY || salary > MAX_SALARY) {
            throw new ValidationException("Salary must be between " + MIN_SALARY + " and " + MAX_SALARY);
        }
    }
    
    /**
     * Validates department field.
     */
    public static void validateDepartment(String department) {
        if (department != null && department.length() > MAX_DEPARTMENT_LENGTH) {
            throw new ValidationException("Department exceeds maximum length of " + MAX_DEPARTMENT_LENGTH + " characters");
        }
    }

    public static HttpHeaders createExcelHeaders(String fileName, byte[] content) {
        HttpHeaders headers = new HttpHeaders();

        // Encode filename for special characters
        String encodedFileName = URLEncoder.encode(fileName, StandardCharsets.UTF_8)
                .replaceAll("\\+", "%20");

        headers.setContentType(MediaType.parseMediaType(TYPE));
        headers.setContentDisposition(ContentDisposition.attachment()
                .filename(encodedFileName)
                .build());
        headers.setCacheControl(CACHE_HEADER);
        headers.setPragma(CACHE);
        headers.set(CONTEXT_TYPE, SNIFF);
        headers.setContentLength(content.length);

        return headers;
    }

    public static void removeMetaData(XSSFWorkbook workbook) {
        // Clear metadata
        POIXMLProperties props = workbook.getProperties();
        props.getCoreProperties().setTitle("");
        props.getCoreProperties().setCreator("");
        props.getExtendedProperties().getUnderlyingProperties().setApplication("");
        props.getCustomProperties().getUnderlyingProperties().getPropertyList().clear();
    }


    public static void headerChecks(List<String> rowData, Map<String, Integer> headerMap, List<String> expectedHeaders) throws BadRequestException {
        for (int i = 0; i < rowData.size(); i++) {
            String cellValue = rowData.get(i) != null ? safeString(rowData.get(i).trim().toLowerCase()) : "";
            headerMap.put(cellValue, i);
            logger.info("Header: [{}] - Index: {}", cellValue, i);
        }

        List<String> missingHeaders = new ArrayList<>();
        for (String expectedHeader : expectedHeaders) {
            if (!headerMap.containsKey(expectedHeader)) {
                missingHeaders.add(expectedHeader);
            }
        }

        if (!missingHeaders.isEmpty()) {
            logger.error("Missing headers: {}", missingHeaders);
            throw new BadRequestException("The following required columns are missing: " + String.join(COMMA, missingHeaders));
        }
    }

    public static String safeString(String value) {
        return ExcelIUtil.sanitize(value).trim();
    }

}

