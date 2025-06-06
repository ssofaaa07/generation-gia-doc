package com.butovetskaia.generationgiadoc.service;

import com.butovetskaia.generationgiadoc.enums.DirectionEducation;
import com.butovetskaia.generationgiadoc.enums.Faculty;
import com.butovetskaia.generationgiadoc.model.DocumentInfo;
import com.butovetskaia.generationgiadoc.model.StudentInfo;
import com.butovetskaia.generationgiadoc.model.InputRequest;
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
import java.util.Arrays;
import java.util.List;

@Service
@RequiredArgsConstructor
public class GenerationDocumentService {
    private static final String TEMPLATE_FILE_NAME = "excel-template";
    private final ExcelService excelService;

    public ResponseEntity<Resource> getExcelTemplate() {
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attchment; filename=\"" + "Данные о ВКР студентов.xlsx\"")
                .header(HttpHeaders.CONTENT_TYPE, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                .body(new ClassPathResource(TEMPLATE_FILE_NAME + ".xlsx"));
    }

    public ResponseEntity<InputStreamResource> generate(InputRequest request) {
        List<StudentInfo> infoStudents = excelService.getStudentInfo(request.getFile(), request.getSchedules());
        DocumentInfo info = DocumentInfo.builder()
                .infoStudents(infoStudents)
                .direction(request.getDirection())
                .faculty(request.getFaculty())
                .formEducation(request.getFormEducation())
                .department(request.getDepartment())
                .qualification(request.getQualification())
                .schedules(request.getSchedules())
                .chairpersonName(request.getChairperson())
                .secretaryName(request.getSecretary())
                .commissionMembers(request.getCommissionMembers())
                .chairpersonAppellateName(request.getChairpersonAppellate())
                .commissionAppellateMembers(request.getCommissionAppellateMembers())
                .declineNames(request.isDeclineNames()).build();

        ByteArrayOutputStream zipOutputStream = DocumentGeneration.generateForProcessOfGia(info);
        ByteArrayInputStream zipInputStream = new ByteArrayInputStream(zipOutputStream.toByteArray());
        String encodedFileName = URLEncoder.encode("Документы ГИА.zip", StandardCharsets.UTF_8).replace("+", "%20");
        return ResponseEntity.ok()
                .header(HttpHeaders.CONTENT_DISPOSITION, "attachment; filename=\"" + encodedFileName + "\"; filename*=UTF-8''" + encodedFileName)
                .header(HttpHeaders.CONTENT_TYPE, MediaType.APPLICATION_OCTET_STREAM_VALUE)
                .contentLength(zipInputStream.available())
                .body(new InputStreamResource(zipInputStream));
    }

    public List<String> getDirection() {
        return DirectionEducation.getStringValues();
    }

    public List<String> getFaculty() {
        return Arrays.stream(Faculty.values()).map(Faculty::getCode).toList();
    }
}
