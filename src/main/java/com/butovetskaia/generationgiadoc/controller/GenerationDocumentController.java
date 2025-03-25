package com.butovetskaia.generationgiadoc.controller;

import com.butovetskaia.generationgiadoc.model.InputRequest;
import com.butovetskaia.generationgiadoc.service.GenerationDocumentService;
import com.fasterxml.jackson.databind.ObjectMapper;
import io.swagger.v3.oas.annotations.Operation;
import jakarta.validation.Valid;
import org.springframework.beans.factory.annotation.Autowired;
import org.springframework.core.io.InputStreamResource;
import org.springframework.core.io.Resource;
import org.springframework.http.ResponseEntity;
import org.springframework.validation.annotation.Validated;
import org.springframework.web.bind.annotation.*;
import org.springframework.web.multipart.MultipartFile;

import java.io.IOException;
import java.util.List;

@RestController
@Validated
public class GenerationDocumentController {
    private final GenerationDocumentService generationDocumentService;

    @Autowired
    public GenerationDocumentController(GenerationDocumentService generationDocumentService) {
        this.generationDocumentService = generationDocumentService;
    }

    @CrossOrigin(origins = "http://localhost:3000")
    @GetMapping("/template")
    @Operation(summary = "Получение шаблона для заполнения информации по студентам")
    public ResponseEntity<Resource> getExcelTemplate() {
        return generationDocumentService.getExcelTemplate();
    }

    @CrossOrigin(origins = "http://localhost:3000")
    @PostMapping(value = "/generate", consumes = "multipart/form-data")
    @Operation(summary = "Генерация пакета документов, сопутствующих процессу ГИА")
    public ResponseEntity<InputStreamResource> generateDocument(
            @RequestPart("file") MultipartFile file,
            @Valid @RequestPart("data") String data) throws IOException {
        final ObjectMapper objectMapper = new ObjectMapper();
        InputRequest request = objectMapper.readValue(data, InputRequest.class);
        request.setFile(file);
        return generationDocumentService.generate(request);
    }

    @CrossOrigin(origins = "http://localhost:3000")
    @GetMapping("/direction")
    @Operation(summary = "Направления обучения для выпадающего списка")
    public ResponseEntity<List<String>> getDirection() {
        return ResponseEntity.ok(generationDocumentService.getDirection());
    }

    @CrossOrigin(origins = "http://localhost:3000")
    @GetMapping("/faculty")
    @Operation(summary = "Факультеты для выпадающего списка")
    public ResponseEntity<List<String>> getFaculty() {
        return ResponseEntity.ok(generationDocumentService.getFaculty());
    }

}
