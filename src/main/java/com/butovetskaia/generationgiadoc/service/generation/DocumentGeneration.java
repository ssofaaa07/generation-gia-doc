package com.butovetskaia.generationgiadoc.service.generation;

import com.butovetskaia.generationgiadoc.model.DateInfo;
import com.butovetskaia.generationgiadoc.model.DocumentInfo;
import com.butovetskaia.generationgiadoc.model.DocumentResultInfo;
import com.butovetskaia.generationgiadoc.service.ExcelService;
import lombok.RequiredArgsConstructor;

import java.io.ByteArrayOutputStream;
import java.time.LocalDate;
import java.util.Comparator;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RequiredArgsConstructor
public abstract class DocumentGeneration {

    protected static final String MARK_LIST_FILE_NAME = "Оценочный лист";
    protected static final String ANNEX_PROTOCOL_FILE_NAME = "Приложение к протоколу ГЭК о защите";
    protected static final String MEETING_PROTOCOL_FILE_NAME = "Протокол заседания ГЭК";
    protected static final String EXCEL_TEMPLATE_RESULT = "Шаблон таблицы с результатами ГИА";
    protected static final String ASSIGNMENT_OF_QUALIFICATION_FILE_NAME = "Приложение о присвоении квалификации выпускникам";
    protected static final String REPORT_FILE_NAME = "Отчет";

    public static ByteArrayOutputStream generateForProcessOfGia(DocumentInfo info) {
        var count = 1;

        MarkListGeneration markListGeneration = new MarkListGeneration();
        AnnexProtocolGeneration annexProtocolGeneration = new AnnexProtocolGeneration(count);
        MeetingProtocolGeneration meetingProtocolGeneration = new MeetingProtocolGeneration(count, false);
        ExcelService excelService = new ExcelService();

        ByteArrayOutputStream zipOutputStream = new ByteArrayOutputStream();
        try (ZipOutputStream zipOut = new ZipOutputStream(zipOutputStream)) {

            var dateExamList = info.schedules().stream().distinct().sorted(Comparator.comparing(DateInfo::getDate)).toList();
            for (var date : dateExamList) {
                ByteArrayOutputStream markListOutputStream = markListGeneration.generationDocument(info, date);
                zipOut.putNextEntry(new ZipEntry(MARK_LIST_FILE_NAME + " " + getDateString(date.getDate()) + ".docx"));
                zipOut.write(markListOutputStream.toByteArray());
                zipOut.closeEntry();

                annexProtocolGeneration.setAnnexNumber(count);
                ByteArrayOutputStream annexProtocolOutputStream = annexProtocolGeneration.generationDocument(info, date);
                zipOut.putNextEntry(new ZipEntry(ANNEX_PROTOCOL_FILE_NAME + " " + getDateString(date.getDate()) + ".docx"));
                zipOut.write(annexProtocolOutputStream.toByteArray());
                zipOut.closeEntry();

                meetingProtocolGeneration.setProtocolNumber(count);
                ByteArrayOutputStream meetingProtocolOutputStream = meetingProtocolGeneration.generationDocument(info, date);
                zipOut.putNextEntry(new ZipEntry(MEETING_PROTOCOL_FILE_NAME + "_" + count + ".docx"));
                zipOut.write(meetingProtocolOutputStream.toByteArray());
                zipOut.closeEntry();

                count++;
            }

            ByteArrayOutputStream excelTemplateForResult = excelService.getTemplateForResult(info);
            zipOut.putNextEntry(new ZipEntry(EXCEL_TEMPLATE_RESULT + ".xlsx"));
            zipOut.write(excelTemplateForResult.toByteArray());
            zipOut.closeEntry();

            return zipOutputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("dgfsjfg");
        }
    }

    public static ByteArrayOutputStream generateForSummingUp(DocumentResultInfo info) {
        AssignmentOfQualificationGeneration assignmentOfQualificationGeneration = new AssignmentOfQualificationGeneration();
        MeetingProtocolGeneration meetingProtocolGeneration = new MeetingProtocolGeneration(info.numberOfProtocol(), true);
        ReportGeneration reportGeneration = new ReportGeneration();

        ByteArrayOutputStream zipOutputStream = new ByteArrayOutputStream();
        try (ZipOutputStream zipOut = new ZipOutputStream(zipOutputStream)) {

            ByteArrayOutputStream meetingProtocolOutputStream = meetingProtocolGeneration.generationDocument(DocumentInfo.of(info), info.date());
            zipOut.putNextEntry(new ZipEntry(MEETING_PROTOCOL_FILE_NAME + "_" + info.numberOfProtocol() + ".docx"));
            zipOut.write(meetingProtocolOutputStream.toByteArray());
            zipOut.closeEntry();

            ByteArrayOutputStream assignmentOfQualificationOutputStream = assignmentOfQualificationGeneration.generationDocument(info, info.date());
            zipOut.putNextEntry(new ZipEntry(ASSIGNMENT_OF_QUALIFICATION_FILE_NAME + ".docx"));
            zipOut.write(assignmentOfQualificationOutputStream.toByteArray());
            zipOut.closeEntry();

            ByteArrayOutputStream reportOutputStream = reportGeneration.generationDocument(info);
            zipOut.putNextEntry(new ZipEntry(REPORT_FILE_NAME + " " + info.direction() + " " + getShortName(info.secretaryName().getMemberName()) + ".docx"));
            zipOut.write(reportOutputStream.toByteArray());
            zipOut.closeEntry();

            return zipOutputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("dgfsjfg");
        }
    }

    public abstract ByteArrayOutputStream generationDocument(DocumentInfo info, DateInfo date);

    public static String getDateString(LocalDate date) {
        var day = String.valueOf(date.getDayOfMonth());
        var month = String.valueOf(date.getMonthValue());
        var year = String.valueOf(date.getYear());

        if (day.length() < 2) {
            day = "0" + day;
        }
        if (month.length() < 2) {
            month = "0" + month;
        }
        return day + "." + month + "." + year;
    }

    public static String getShortName(String name) {
        String[] parts = name.trim().split("\\s+");

        if (parts.length < 2) {
            throw new IllegalArgumentException("Full name must contain at least two parts: last name and first name");
        }

        String lastName = parts[0];
        String firstNameInitial = parts[1].charAt(0) + ".";

        if (parts.length >= 3) {
            String patronymicInitial = parts[2].charAt(0) + ".";
            return lastName + " " + firstNameInitial + patronymicInitial;
        } else {
            return lastName + " " + firstNameInitial;
        }
    }
}
