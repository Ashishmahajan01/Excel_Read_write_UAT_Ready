package com.excelDemo.excel.dto;

import com.fasterxml.jackson.annotation.JsonInclude;
import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import java.time.LocalDateTime;
import java.util.List;

@Data
@AllArgsConstructor
@NoArgsConstructor
public class FileResponse {
    public String fileName;
    public Integer size;
    public LocalDateTime uploadDate;
    public String status;
    public List<String> errors;
    public Integer recordsProcessed;
}
