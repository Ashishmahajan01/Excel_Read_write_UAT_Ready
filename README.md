## Excel Demo – Secure Excel Read/Write with Spring Boot

This project is a **Spring Boot demo** that shows how to **upload, validate, sanitize, store, and download Excel data** using **Apache POI** and **MySQL**.  

In addition to basic Excel processing, it focuses on **security and compliance** so the APIs are suitable for **UAT and penetration testing**:

- **Sanitized strings**: every text value coming from Excel is passed through a `sanitize` / `safeString` layer (`ExcelIUtil.sanitize`) that:
  - Detects and blocks script/XSS payloads using multiple regex patterns (script tags, inline event handlers, `javascript:` URIs, dangerous HTML tags, etc.).
  - Neutralizes Excel formula injection by prefixing risky values (starting with `=`, `+`, `-`, `@`, `<`, `>` etc.) with a safe `'` character.
  - Trims and normalizes the content before it is stored or returned.
  - Throws `ValidationException` (HTTP 400) instead of generic exceptions (HTTP 500) for better pen-test compliance.
- **Excel metadata removal**: before sending any Excel file back to the client, the code calls `ExcelIUtil.removeMetaData(workbook)` to:
  - Strip author, application, and custom properties from the workbook.
  - Reduce the risk of leaking internal environment details or user information through document properties.
- **Magic byte validation**: file type is verified using both Content-Type header and file signature (magic bytes) to prevent attackers from uploading malicious files with fake Content-Type headers.
- **Row limit validation**: enforces a maximum of 10,000 rows per file to prevent DoS attacks from extremely large files.
- **Field-level validation**: comprehensive validation for all fields:
  - **Username**: non-empty, max 100 characters
  - **Email**: valid email format (regex), max 255 characters
  - **Age**: range validation (0-150)
  - **Salary**: range validation (0 - 10,000,000)
  - **Department**: max 100 characters (optional field)
- **Proper HTTP status codes**: validation errors return `400 Bad Request` instead of `500 Internal Server Error` for better security posture.
- **Security headers**: proper `X-Content-Type-Options: nosniff` header with correct casing for pen-test compliance.

These measures help the application **pass UAT and security reviews**, because:

- Uploaded content is validated and sanitized to prevent **XSS / script injection / formula injection** scenarios when data is later opened or rendered.
- Generated Excel reports do not expose **hidden metadata** that security and pen-test teams often flag.
- File uploads are validated at multiple layers (Content-Type + magic bytes) to prevent malicious file uploads.
- Resource limits prevent DoS attacks from oversized files.
- Field validation ensures data integrity and prevents injection attacks.

It exposes simple REST APIs that let you:

- **Upload an Excel file** with sanitized user data and store it in a database.
- **Validate headers and row values** with clear error messages.
- **Download all users** back as a formatted, metadata-clean Excel report.

You can use this README as the main description when you publish the repository on GitHub.

---

### Tech stack

- **Language**: Java 21  
- **Framework**: Spring Boot 3.5.5  
- **Build tool**: Maven  
- **Database**: MySQL  
- **Excel library**: Apache POI (`poi`, `poi-ooxml`)  
- **ORM**: Spring Data JPA  
- **Testing**: Spring Boot Test / JUnit 5  

---

### Project highlights

- **Excel upload API**  
  - Endpoint: `POST /api/excel/upload`  
  - Accepts `.xlsx` files via `multipart/form-data` (CSV files are not supported).  
  - **Multi-layer validation**:
    - File format validation: Content-Type header + magic byte (file signature) verification
    - File size limit: 10MB maximum (configurable in `application.properties`)
    - Row limit: maximum 10,000 rows per file (prevents DoS attacks)
    - Header row validation: matches expected columns exactly
    - Field-level validation: length limits, ranges, email format
    - Sanitization: all text fields sanitized for XSS/script injection prevention
  - Each row is converted to a `UserData` entity with validation.
  - Stores valid rows in MySQL and returns a detailed summary (`FileResponse`) with:
    - File name and size  
    - Upload time  
    - Status message (success / partial success / failed)  
    - Number of records processed  
    - Per‑row error messages, if any  

- **Excel download API**  
  - Endpoint: `GET /api/excel/users`  
  - Reads all `UserData` records from the database.  
  - Generates a clean, auto-sized Excel report with:
    - Header styling  
    - Number formatting  
    - Date formatting for `createdAt`  
  - Returns the Excel file as a binary response so it can be downloaded from a browser or Postman.

- **Robust error handling**  
  - Custom exceptions:
    - `ValidationException` - Input validation errors (HTTP 400)
    - `InvalidFileFormatException` - Invalid file format (HTTP 400)
    - `ExcelProcessingException` - Excel processing errors (HTTP 400)
    - `FileUploadException` - File upload errors (HTTP 500)
  - Centralized `GlobalExceptionHandler` (Spring `@ControllerAdvice`) to convert errors into consistent JSON responses.
  - Proper HTTP status codes: validation errors return 400 Bad Request, not 500 Internal Server Error.

---

### Excel format expected by the upload API

The upload API expects the first sheet to have this **header row** (case-insensitive, order as below):

| Column name | Description              | Example           |
|------------|--------------------------|-------------------|
| `username` | User name                | `john.doe`        |
| `email`    | Email address            | `john@demo.com`   |
| `age`      | Age (integer)            | `30`              |
| `department` | Department / team name | `IT`              |
| `salary`   | Salary (decimal/number)  | `50000`           |
| `is_active`| Active flag (true/false) | `true` or `false` |

The service:

- Validates the header row (case-insensitive, exact column match required).
- Validates each data row with field-level constraints:
  - **Username**: Required, 1-100 characters, sanitized
  - **Email**: Required, valid email format, max 255 characters, sanitized
  - **Age**: Required, integer between 0-150
  - **Department**: Optional, max 100 characters, sanitized
  - **Salary**: Required, decimal between 0-10,000,000
  - **Is Active**: Required, boolean (true/false)
- Sanitizes all text fields to prevent XSS and script injection.
- Collects per-row errors and still processes valid rows (partial success is supported).
- Enforces row limit (max 10,000 rows) to prevent DoS attacks.

---

### API endpoints

- **Upload Excel**
  - **Method**: `POST`  
  - **URL**: `http://localhost:8600/api/excel/upload`  
  - **Headers**: `Content-Type: multipart/form-data`  
  - **Body**: form-data with key `file` and value as the `.xlsx` file.

  **Sample response (success)**:

  ```json
  {
    "fileName": "users.xlsx",
    "fileSize": 10240,
    "uploadTime": "2026-01-27T10:15:30.123",
    "status": "Success - 10 records processed successfully",
    "errors": null,
    "recordsProcessed": 10
  }
  ```

  **Sample response (validation error)**:

  ```json
  {
    "fileName": "users.xlsx",
    "fileSize": 10240,
    "uploadTime": "2026-01-27T10:15:30.123",
    "status": "Failed",
    "errors": [
      "File contains too many rows (15000). Maximum allowed: 10000"
    ],
    "recordsProcessed": 0
  }
  ```

  **Sample response (partial success)**:

  ```json
  {
    "fileName": "users.xlsx",
    "fileSize": 10240,
    "uploadTime": "2026-01-27T10:15:30.123",
    "status": "Partial success - 8 records processed with 2 errors",
    "errors": [
      "Error processing row 5: Invalid email format: invalid-email",
      "Error processing row 10: Age must be between 0 and 150"
    ],
    "recordsProcessed": 8
  }
  ```

- **Download all users as Excel**
  - **Method**: `GET`  
  - **URL**: `http://localhost:8600/api/excel/users`  
  - **Response**: Excel file (e.g. `User_Demo.xlsx`) as binary (`application/vnd.openxmlformats-officedocument.spreadsheetml.sheet`).

---

### How to run the project locally

- **Prerequisites**
  - Java 21 installed
  - Maven installed (or use the included `mvnw`/`mvnw.cmd`)
  - MySQL running locally

- **1. Configure the database**

  Update `src/main/resources/application.properties` with your own MySQL settings:

  ```properties
  spring.datasource.url=jdbc:mysql://localhost:3306/<your_db>?serverTimezone=UTC&allowPublicKeyRetrieval=true&useSSL=false
  spring.datasource.username=<your_username>
  spring.datasource.password=<your_password>

  spring.jpa.hibernate.ddl-auto=update
  spring.jpa.open-in-view=false
  server.port=8600
  ```

  Make sure the database (`<your_db>`) exists before starting the application.

- **2. Build and run**

  From the project root:

  ```bash
  # Build
  mvn clean install

  # Run
  mvn spring-boot:run
  ```

  The application will start on `http://localhost:8600`.

---

### Testing the APIs with Postman

- A ready-made Postman collection is included: **`Excel_Demo.postman_collection.json`**.
- Import this file into Postman.
- It contains:
  - A **POST /api/excel/upload** request where you can attach your Excel file.
  - A **GET /api/excel/users** request to download the generated Excel.

This makes it easy for others to quickly test your API without writing any extra code.

---

### Running tests

To run the test suite:

```bash
mvn test
```

The project currently includes a basic Spring Boot context load test, and you can easily extend it with more tests for the upload/download services.

---

### Security features summary

This project implements multiple security layers suitable for UAT and penetration testing:

| Security Feature | Implementation | Purpose |
|-----------------|----------------|---------|
| **Input Sanitization** | `ExcelIUtil.sanitize()` | Prevents XSS, script injection, formula injection |
| **Magic Byte Validation** | `ExcelIUtil.hasExcelFormat()` | Verifies file type using file signature, not just Content-Type |
| **Row Limit** | `MAX_ROWS_LIMIT = 10000` | Prevents DoS attacks from oversized files |
| **Field Validation** | `validateEmail()`, `validateUsername()`, etc. | Ensures data integrity and prevents injection |
| **Metadata Removal** | `ExcelIUtil.removeMetaData()` | Strips sensitive metadata from generated Excel files |
| **Proper HTTP Status Codes** | `ValidationException` → 400 | Returns 400 for validation errors, not 500 |
| **Security Headers** | `X-Content-Type-Options: nosniff` | Prevents MIME type sniffing attacks |

---

### Project structure (high level)

- `pom.xml` – Maven configuration and dependencies (Spring Boot, JPA, MySQL, Apache POI, Lombok, tests).  
- `ExcelDemoApplication.java` – Main Spring Boot application entry point.  
- `controller/ExcelUploadController.java` – REST APIs for uploading and downloading Excel files.  
- `service/ExcelUploadService.java` – Reads and validates Excel upload, maps rows to `UserData`, saves to DB. Includes row limit validation and field-level validation.  
- `service/ExcelDownloadService.java` – Builds an Excel report from data stored in the DB. Removes metadata before sending.  
- `model/UserData.java` – JPA entity representing a user row.  
- `repository/UserDataRepository.java` – Spring Data JPA repository.  
- `util/ExcelIUtil.java` – Core utility class containing:
  - Sanitization methods (`sanitize()`, `isScriptInjection()`)
  - File validation (`hasExcelFormat()`, `validateExcelSignature()`)
  - Field validation methods (`validateEmail()`, `validateUsername()`, etc.)
  - Metadata removal (`removeMetaData()`)
  - Security header creation (`createExcelHeaders()`)
- `exception/*` – Custom exceptions:
  - `ValidationException` – Input validation errors (HTTP 400)
  - `InvalidFileFormatException` – Invalid file format (HTTP 400)
  - `ExcelProcessingException` – Excel processing errors (HTTP 400)
  - `FileUploadException` – File upload errors (HTTP 500)
  - `GlobalExceptionHandler` – Centralized exception handling
- `src/test/java/.../ExcelDemoApplicationTests.java` – Basic Spring Boot test class.  

---

### How others can use this repo

- As a **reference implementation** for integrating Excel import/export into Spring Boot with security best practices.  
- As a **starting point** to build more complex reporting or bulk data upload features.  
- As a **learning project** to understand how to:
  - Work with Apache POI for Excel processing
  - Implement input sanitization to prevent XSS and script injection
  - Validate file uploads using magic bytes (file signatures)
  - Apply field-level validation (lengths, ranges, formats)
  - Handle errors properly with appropriate HTTP status codes
  - Remove sensitive metadata from generated files
  - Implement DoS protection with resource limits
- As a **security reference** for:
  - Penetration testing preparation
  - UAT (User Acceptance Testing) compliance
  - Security audit preparation
  - Understanding common Excel-related vulnerabilities and mitigations

### Key security takeaways

This project demonstrates how to:

1. **Sanitize user input** before storing in the database to prevent XSS and injection attacks.
2. **Validate file types** using both Content-Type headers and magic bytes (file signatures).
3. **Enforce resource limits** (row counts, file sizes) to prevent DoS attacks.
4. **Validate field constraints** (lengths, ranges, formats) to ensure data integrity.
5. **Return proper HTTP status codes** (400 for validation errors, not 500).
6. **Remove metadata** from generated files to prevent information leakage.
7. **Use security headers** (`X-Content-Type-Options`) to prevent MIME type sniffing.

Feel free to fork this repository and adapt it to your own use cases.

