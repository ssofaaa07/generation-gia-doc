package com.butovetskaia.generationgiadoc.service;

import com.butovetskaia.generationgiadoc.model.InputResultRequest;
import com.butovetskaia.generationgiadoc.service.generation.DocumentGeneration;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.ClassPathResource;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.HttpHeaders;
import org.springframework.http.MediaType;
import org.springframework.http.ResponseEntity;
import org.springframework.stereotype.Service;

import java.io.ByteArrayInputStream;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.net.URLEncoder;
import java.nio.charset.StandardCharsets;

@Service
@RequiredArgsConstructor
public class ResultService {

    private static final String TEMPLATE_FILE_NAME = "excel-template-mark";
    private final ExcelService excelService;

    public ResponseEntity<Resource> getExcelTemplate() {
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attchment; filename=\"" + TEMPLATE_FILE_NAME + ".xlsx\"")
                .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(new ClassPathResource(TEMPLATE_FILE_NAME + ".xlsx"));
    }

    public ResponseEntity<InputStreamResource> generate(InputResultRequest request) {
        var info = excelService.getDocumentInfo(request.getFile(), request.getSchedule(), request.isDeclineNames());

        ByteArrayOutputStream zipOutputStream = DocumentGeneration.generateForSummingUp(info);
        ByteArrayInputStream zipInputStream = new ByteArrayInputStream(zipOutputStream.toByteArray());
        String encodedFileName = URLEncoder.encode("Результаты ГИА.zip", StandardCharsets.UTF_8).replace("+", "%20");
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + encodedFileName + "\"; filename*=UTF-8''" + encodedFileName)
                .header(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_OCTET_STREAM_VALUE)
                .contentLength(zipInputStream.available())
                .body(new InputStreamResource(zipInputStream));
    }
}
