package com.butovetskaia.generationgiadoc.controller;

import com.butovetskaia.generationgiadoc.model.InputResultRequest;
import com.butovetskaia.generationgiadoc.service.ResultService;
import com.fasterxml.jackson.databind.DeserializationFeature;
import com.fasterxml.jackson.databind.ObjectMapper;
import com.fasterxml.jackson.databind.SerializationFeature;
import com.fasterxml.jackson.datatype.jsr310.JavaTimeModule;
import io.swagger.v3.oas.annotations.Operation;
import jakarta.validation.Valid;
import lombok.RequiredArgsConstructor;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;

@RestController
@RequestMapping("/results")
@Validated
@RequiredArgsConstructor
public class ResultController {
    private final ResultService resultService;

    @CrossOrigin(origins = "http://localhost:3000")
    @GetMapping("/template")
    @Operation(summary = "Получение шаблона для заполнения информации о результатах проведения ГИА")
    public ResponseEntity<Resource> getExcelTemplate() {
        return resultService.getExcelTemplate();
    }

    @CrossOrigin(origins = "http://localhost:3000")
    @PostMapping(value = "/documents/generate", consumes = "multipart/form-data")
    @Operation(summary = "Генерация документов о проведении ГИА для отчетности")
    public ResponseEntity<InputStreamResource> generateDocument(
            @RequestPart("file") MultipartFile file,
            @Valid @RequestPart("data") String data) throws IOException {
        final ObjectMapper objectMapper = new ObjectMapper();
        objectMapper.registerModule(new JavaTimeModule());
        objectMapper.disable(SerializationFeature.WRITE_DATES_AS_TIMESTAMPS);
        objectMapper.disable(DeserializationFeature.ADJUST_DATES_TO_CONTEXT_TIME_ZONE);
        InputResultRequest request = objectMapper.readValue(data, InputResultRequest.class);
        request.setFile(file);
        return resultService.generate(request);
    }
}
