package com.butovetskaia.generationgiadoc.service;

import com.aspose.cells.Cells;
import com.aspose.cells.SaveFormat;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.butovetskaia.generationgiadoc.enums.Recommendation;
import com.butovetskaia.generationgiadoc.enums.Theme;
import com.butovetskaia.generationgiadoc.model.*;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.time.LocalDate;
import java.util.ArrayList;
import java.util.Comparator;
import java.util.List;

import static org.apache.commons.lang3.StringUtils.isNotBlank;

@Service
@Slf4j
public class ExcelService {

    public List<StudentInfo> getStudentInfo(MultipartFile file, List<DateInfo> schedules) {
        try {
            Workbook workbook = new Workbook(file.getInputStream());

            List<StudentInfo> infoStudents = new ArrayList<>();
            var sortedSchedules = schedules.stream().sorted(Comparator.comparing(DateInfo::getDate)).toArray();

            for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
                Worksheet worksheet = workbook.getWorksheets().get(i);
                Cells cells = worksheet.getCells();

                infoStudents.addAll(readStudentInfo(cells, (DateInfo) sortedSchedules[i]));
            }

            return infoStudents;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private List<StudentInfo> readStudentInfo(Cells cells, DateInfo dateExam) {
        List<StudentInfo> infoStudents = new ArrayList<>();

        for (int row = 2; row < cells.getMaxDataRow() + 1; row++) {

            var themeType = getTheme(cells, row);
            var recommendations = getRecommendations(cells, row);

            if (isNotBlank(cells.get(row, 2).getStringValue().trim())) {
                StudentInfo info = StudentInfo.builder()
                        .dateExam(dateExam)
                        .name(cells.get(row, 1).getStringValue().trim())
                        .themeName(cells.get(row, 2).getStringValue().trim())
                        .supervisorName(cells.get(row, 3).getStringValue().trim())
                        .supervisorMark(cells.get(row, 4).getStringValue().trim())
                        .themeType(themeType)
                        .recommendations(recommendations)
                        .markOne(cells.get(row, 11).getIntValue())
                        .markTwo(cells.get(row, 12).getIntValue()).build();
                infoStudents.add(info);
            }
        }
        return infoStudents;
    }

    private Theme getTheme(Cells cells, int row) {
        if (cells.get(row, 6).getStringValue().trim() != null && !cells.get(row, 6).getStringValue().trim().isEmpty()) {
            return Theme.ENTERPRISE_THEME;
        } else if (cells.get(row, 7).getStringValue().trim() != null && !cells.get(row, 7).getStringValue().trim().isEmpty()) {
            return Theme.RESEARCH_THEME;
        }
        return Theme.STUDENT_THEME;
    }

    private List<Recommendation> getRecommendations(Cells cells, int row) {
        List<Recommendation> recommendations = new ArrayList<>();
        if (cells.get(row, 8).getStringValue().trim() != null && !cells.get(row, 8).getStringValue().trim().isEmpty()) {
            recommendations.add(Recommendation.PUBLICATION_RECOMMENDATION);
        }
        if (cells.get(row, 9).getStringValue().trim() != null && !cells.get(row, 9).getStringValue().trim().isEmpty()) {
            recommendations.add(Recommendation.IMPLEMENTATION_RECOMMENDATION);
        }
        if (cells.get(row, 10).getStringValue().trim() != null && !cells.get(row, 10).getStringValue().trim().isEmpty()) {
            recommendations.add(Recommendation.IMPLEMENTED);
        }
        return recommendations;
    }

    private List<StudentResultInfo> readStudentResultInfo(Cells cells) {
        List<StudentResultInfo> infoStudents = new ArrayList<>();
        for (int row = 2; row < cells.getMaxDataRow() + 1; row++) {
            if (isNotBlank(cells.get(row, 2).getStringValue().trim())) {
                StudentResultInfo info = StudentResultInfo.builder()
                        .name(cells.get(row, 1).getStringValue().trim())
                        .themeName(cells.get(row, 2).getStringValue().trim())
                        .supervisorName(cells.get(row, 3).getStringValue().trim())
                        .mark(cells.get(row, 4).getIntValue())
                        .isRed(isNotBlank(cells.get(row, 5).getStringValue().trim()) || cells.get(row, 5).getStringValue().trim().equals("+"))
                        .isBest(isNotBlank(cells.get(row, 6).getStringValue().trim()) || cells.get(row, 6).getStringValue().trim().equals("+")).build();

                infoStudents.add(info);
            }
        }
        return infoStudents;
    }

    public DocumentResultInfo getDocumentInfo(MultipartFile file, DateInfo dateCommission) {
        try {
            var builder = DocumentResultInfo.builder();
            Workbook workbook = new Workbook(file.getInputStream());

            Worksheet worksheet = workbook.getWorksheets().get(0);
            Cells cells = worksheet.getCells();

            builder.infoStudents(readStudentResultInfo(cells));

            worksheet = workbook.getWorksheets().get(1);

            getInfoFromWorksheet(worksheet, builder, dateCommission);

            return builder.build();
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private void getInfoFromWorksheet(Worksheet worksheet, DocumentResultInfo.DocumentResultInfoBuilder builder, DateInfo dateCommission) {
        var row = 0;
        builder.faculty(worksheet.getCells().get(row++, 1).getStringValue().trim());
        builder.direction(worksheet.getCells().get(row++, 1).getStringValue().trim());
        builder.department(worksheet.getCells().get(row++, 1).getStringValue().trim());
        builder.qualification(worksheet.getCells().get(row++, 1).getStringValue().trim());
        builder.formEducation(worksheet.getCells().get(row++, 1).getStringValue().trim());

        row++;

        var col = 1;
        LocalDate orderDate = null;
        int orderNumber = 0;
        while (isNotBlank(worksheet.getCells().get(row, col).getStringValue().trim())) {
            col++;
            orderNumber = worksheet.getCells().get(row + 4, 1).getIntValue();
            orderDate = LocalDate.parse(worksheet.getCells().get(row + 5, 1).getStringValue().trim());
        }

        dateCommission.setOrderNumber(orderNumber);
        dateCommission.setOrderDate(orderDate);

        builder.numberOfProtocol(col)
                .date(dateCommission);

        row = row + 8;

        builder.chairpersonName(new CommissionMemberInfo(
            worksheet.getCells().get(row++, 1).getStringValue().trim(),
            worksheet.getCells().get(row++, 1).getStringValue().trim(),
            worksheet.getCells().get(row++, 1).getIntValue(),
            LocalDate.parse(worksheet.getCells().get(row++, 1).getStringValue().trim())
                        ));

        row++;
        row++;

        builder.secretaryName(new CommissionMemberInfo(
                worksheet.getCells().get(row++, 1).getStringValue().trim(),
                worksheet.getCells().get(row++, 1).getStringValue().trim(),
                worksheet.getCells().get(row++, 1).getIntValue(),
                LocalDate.parse(worksheet.getCells().get(row++, 1).getStringValue().trim())
        ));

        row++;
        row++;

        col = 1;
        var roww = 0;
        List<CommissionMemberInfo> members = new ArrayList<>();
        while (isNotBlank(worksheet.getCells().get(row, col).getStringValue().trim())) {
            roww = row;
            members.add(new CommissionMemberInfo(
                worksheet.getCells().get(roww++, col).getStringValue().trim(),
                worksheet.getCells().get(roww++, col).getStringValue().trim(),
                worksheet.getCells().get(roww++, col).getIntValue(),
                LocalDate.parse(worksheet.getCells().get(roww++, col).getStringValue().trim())
            ));
            col++;
        }

        builder.commissionMembers(members);
        row = ++roww;
        row++;

        builder.chairpersonAppellateName(new CommissionMemberInfo(
                worksheet.getCells().get(row++, 1).getStringValue().trim(),
                worksheet.getCells().get(row++, 1).getStringValue().trim(),
                worksheet.getCells().get(row++, 1).getIntValue(),
                LocalDate.parse(worksheet.getCells().get(row++, 1).getStringValue().trim())
        ));

        row++;
        row++;

        col = 1;
        List<CommissionMemberInfo> appellateMembers = new ArrayList<>();
        while (isNotBlank(worksheet.getCells().get(row, col).getStringValue().trim())) {
            roww = row;
            appellateMembers.add(new CommissionMemberInfo(
                    worksheet.getCells().get(roww++, col).getStringValue().trim(),
                    worksheet.getCells().get(roww++, col).getStringValue().trim(),
                    worksheet.getCells().get(roww++, col).getIntValue(),
                    LocalDate.parse(worksheet.getCells().get(roww++, col).getStringValue().trim())
            ));
            col++;
        }

        builder.commissionAppellateMembers(members);

        row = ++roww;
        row++;

        builder.countStudentTheme(worksheet.getCells().get(row++, 1).getIntValue());
        builder.countEnterpriseTheme(worksheet.getCells().get(row++, 1).getIntValue());
        builder.countResearchTheme(worksheet.getCells().get(row++, 1).getIntValue());

        row++;
        row++;

        builder.countPublicationRecommendation(worksheet.getCells().get(row++, 1).getIntValue());
        builder.countImplementationRecommendation(worksheet.getCells().get(row++, 1).getIntValue());
        builder.countImplemented(worksheet.getCells().get(row++, 1).getIntValue());
    }

    public ByteArrayOutputStream getTemplateForResult(DocumentInfo info) {

        try (InputStream templateStream = getClass().getClassLoader()
                .getResourceAsStream("excel-template-mark.xlsx");

             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            Workbook workbook = new Workbook(templateStream);
            Worksheet worksheet = workbook.getWorksheets().get(0);

            var sortedStudents = info.infoStudents().stream().sorted(Comparator.comparing(StudentInfo::getName)).toList();

            var style = workbook.createStyle();
            style.setTextWrapped(true);

            int rowNum = 2;
            int count = 0;
            for (StudentInfo student : sortedStudents) {
                worksheet.getCells().get(rowNum, 0).putValue(++count);
                worksheet.getCells().get(rowNum, 1).putValue(student.getName());
                worksheet.getCells().get(rowNum, 2).setStyle(style);
                worksheet.getCells().get(rowNum, 2).putValue(student.getThemeName());
                worksheet.getCells().get(rowNum, 3).putValue(student.getSupervisorName());

                rowNum++;
            }

//            worksheet.autoFitColumns();

            fillInfo(workbook, info);

            workbook.save(outputStream, SaveFormat.XLSX);
            return outputStream;
        } catch (Exception e) {
            throw new RuntimeException("Ошибка при заполнении шаблона Excel", e);
        }
    }

    private void fillInfo(Workbook workbook, DocumentInfo info) {
        Worksheet worksheet = workbook.getWorksheets().get(1);

        var row = 0;
        worksheet.getCells().get(row++, 1).putValue(info.faculty());
        worksheet.getCells().get(row++, 1).putValue(info.direction());
        worksheet.getCells().get(row++, 1).putValue(info.department());
        worksheet.getCells().get(row++, 1).putValue(info.qualification());
        worksheet.getCells().get(row++, 1).putValue(info.formEducation());

        row++;

        var col = 1;
        var roww = 0;
        for (var date : info.schedules()) {
            roww = row;
            worksheet.getCells().get(roww++, col).putValue(col);
            worksheet.getCells().get(roww++, col).putValue(date.getDate().toString());
            worksheet.getCells().get(roww++, col).putValue(date.getStart().toString());
            worksheet.getCells().get(roww++, col).putValue(date.getEnd().toString());
            worksheet.getCells().get(roww++, col).putValue(date.getOrderNumber());
            worksheet.getCells().get(roww++, col).putValue(date.getOrderDate().toString());

            col++;
        }

        row = ++roww;
        row++;

        worksheet.getCells().get(row++, 1).putValue(info.chairpersonName().getMemberName());
        worksheet.getCells().get(row++, 1).putValue(info.chairpersonName().getMemberPost());
        worksheet.getCells().get(row++, 1).putValue(info.chairpersonName().getOrderNumber());
        worksheet.getCells().get(row++, 1).putValue(info.chairpersonName().getOrderDate().toString());

        row++;
        row++;

        worksheet.getCells().get(row++, 1).putValue(info.secretaryName().getMemberName());
        worksheet.getCells().get(row++, 1).putValue(info.secretaryName().getMemberPost());
        worksheet.getCells().get(row++, 1).putValue(info.secretaryName().getOrderNumber());
        worksheet.getCells().get(row++, 1).putValue(info.secretaryName().getOrderDate().toString());

        row++;

        col = 1;
        for (var member : info.commissionMembers()) {
            roww = row;
            worksheet.getCells().get(roww++, col).putValue(col);
            worksheet.getCells().get(roww++, col).putValue(member.getMemberName());
            worksheet.getCells().get(roww++, col).putValue(member.getMemberPost());
            worksheet.getCells().get(roww++, col).putValue(member.getOrderNumber());
            worksheet.getCells().get(roww++, col).putValue(member.getOrderDate().toString());

            col++;
        }

        row = ++roww;
        row++;

        worksheet.getCells().get(row++, 1).putValue(info.chairpersonAppellateName().getMemberName());
        worksheet.getCells().get(row++, 1).putValue(info.chairpersonAppellateName().getMemberPost());
        worksheet.getCells().get(row++, 1).putValue(info.chairpersonAppellateName().getOrderNumber());
        worksheet.getCells().get(row++, 1).putValue(info.chairpersonAppellateName().getOrderDate().toString());

        row++;

        col = 1;
        for (var member : info.commissionAppellateMembers()) {
            roww = row;
            worksheet.getCells().get(roww++, col).putValue(col);
            worksheet.getCells().get(roww++, col).putValue(member.getMemberName());
            worksheet.getCells().get(roww++, col).putValue(member.getMemberPost());
            worksheet.getCells().get(roww++, col).putValue(member.getOrderNumber());
            worksheet.getCells().get(roww++, col).putValue(member.getOrderDate().toString());

            col++;
        }

        row = ++roww;
        row++;

        var countStudentTheme = info.infoStudents().stream().filter(it -> it.getThemeType().isStudentTheme()).count();
        var countEnterpriseTheme = info.infoStudents().stream().filter(it -> it.getThemeType().isEnterpriseTheme()).count();
        var countResearchTheme = info.infoStudents().stream().filter(it -> it.getThemeType().isResearchTheme()).count();

        worksheet.getCells().get(row++, 1).putValue(countStudentTheme);
        worksheet.getCells().get(row++, 1).putValue(countEnterpriseTheme);
        worksheet.getCells().get(row++, 1).putValue(countResearchTheme);

        row++;
        row++;

        var countPublicationRecommendation = info.infoStudents().stream().filter(it -> it.getRecommendations().contains(Recommendation.PUBLICATION_RECOMMENDATION)).count();
        var countImplementationRecommendation = info.infoStudents().stream().filter(it -> it.getRecommendations().contains(Recommendation.IMPLEMENTATION_RECOMMENDATION)).count();
        var countImplemented = info.infoStudents().stream().filter(it -> it.getRecommendations().contains(Recommendation.IMPLEMENTED)).count();

        worksheet.getCells().get(row++, 1).putValue(countPublicationRecommendation);
        worksheet.getCells().get(row++, 1).putValue(countImplementationRecommendation);
        worksheet.getCells().get(row++, 1).putValue(countImplemented);
    }
}
