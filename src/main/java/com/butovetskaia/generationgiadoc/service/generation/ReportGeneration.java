package com.butovetskaia.generationgiadoc.service.generation;

import com.aspose.words.*;
import com.butovetskaia.generationgiadoc.model.DocumentResultInfo;
import com.butovetskaia.generationgiadoc.model.StudentResultInfo;
import lombok.AllArgsConstructor;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

import static com.butovetskaia.generationgiadoc.service.generation.DocumentGeneration.getDateString;
import static com.butovetskaia.generationgiadoc.service.generation.DocumentGeneration.getShortName;

@Setter
@Slf4j
@AllArgsConstructor
public class ReportGeneration {

    public ByteArrayOutputStream generationDocument(DocumentResultInfo info) {
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            InputStream inputStream = getClass().getClassLoader().getResourceAsStream(DocumentGeneration.REPORT_FILE_NAME + "-template.docx");
            Document doc = new Document(inputStream);

            doc.getRange().replace("{{faculty}}", info.faculty(), new FindReplaceOptions());
            doc.getRange().replace("{{department}}", info.department(), new FindReplaceOptions());
            doc.getRange().replace("{{direction}}", info.direction(), new FindReplaceOptions());
            doc.getRange().replace("{{yearNow}}", String.valueOf(info.date().getDate().getYear()), new FindReplaceOptions());
            doc.getRange().replace("{{yearPast}}", String.valueOf(info.date().getDate().getYear() - 1), new FindReplaceOptions());

            doc.getRange().replace("{{day_of_date_chairperson}}", String.valueOf(info.chairpersonName().getOrderDate().getDayOfMonth()), new FindReplaceOptions());
            doc.getRange().replace("{{month_of_date_chairperson}}", getMonth(info.chairpersonName().getOrderDate().getMonthValue()), new FindReplaceOptions());
            doc.getRange().replace("{{year_of_date_chairperson}}", String.valueOf(info.chairpersonName().getOrderDate().getYear()), new FindReplaceOptions());
            doc.getRange().replace("{{chairperson_name}}", info.chairpersonName().getMemberName(), new FindReplaceOptions());
            doc.getRange().replace("{{chairperson_post}}", info.chairpersonName().getMemberPost(), new FindReplaceOptions());

            doc.getRange().replace("{{date_of_commission_members}}", getDateString(info.commissionMembers().getFirst().getOrderDate()), new FindReplaceOptions());
            doc.getRange().replace("{{number_of_commission_members}}", String.valueOf(info.commissionMembers().getFirst().getOrderNumber()), new FindReplaceOptions());

            //TODO вывод членов комиссии

            doc.getRange().replace("{{date_of_appellate_chairperson}}", getDateString(info.chairpersonAppellateName().getOrderDate()), new FindReplaceOptions());
            doc.getRange().replace("{{number_of_appellate_chairperson}}", String.valueOf(info.chairpersonAppellateName().getOrderNumber()), new FindReplaceOptions());
            doc.getRange().replace("{{appellate_chairperson_name}}", getInverseName(info.chairpersonAppellateName().getMemberName()), new FindReplaceOptions());
            doc.getRange().replace("{{appellate_chairperson_post}}", info.chairpersonAppellateName().getMemberPost(), new FindReplaceOptions());

            //TODO вывод членов appellate комиссии

            doc.getRange().replace("{{date_of_secretary}}", getDateString(info.secretaryName().getOrderDate()), new FindReplaceOptions());
            doc.getRange().replace("{{number_of_secretary}}", String.valueOf(info.secretaryName().getOrderNumber()), new FindReplaceOptions());
            doc.getRange().replace("{{secretary_name}}", info.secretaryName().getMemberName(), new FindReplaceOptions());
            doc.getRange().replace("{{secretary_post}}", info.secretaryName().getMemberPost(), new FindReplaceOptions());

            doc.getRange().replace("{{date_of_schedules}}", getDateString(info.date().getOrderDate()), new FindReplaceOptions());
            doc.getRange().replace("{{number_of_schedules}}", String.valueOf(info.date().getOrderNumber()), new FindReplaceOptions());

            doc.getRange().replace("{{count_students_theme}}", String.valueOf(info.countStudentTheme()), new FindReplaceOptions());
            doc.getRange().replace("{{count_enterprise_theme}}", String.valueOf(info.countEnterpriseTheme()), new FindReplaceOptions());
            doc.getRange().replace("{{count_research_theme}}", String.valueOf(info.countResearchTheme()), new FindReplaceOptions());
            doc.getRange().replace("{{count_publication_recommendation}}", String.valueOf(info.countPublicationRecommendation()), new FindReplaceOptions());
            doc.getRange().replace("{{count_implementation_recommendation}}", String.valueOf(info.countImplementationRecommendation()), new FindReplaceOptions());
            doc.getRange().replace("{{count_implemented}}", String.valueOf(info.countImplemented()), new FindReplaceOptions());

            doc.getRange().replace("{{chairperson}}", getShortName(info.chairpersonName().getMemberName()), new FindReplaceOptions());
            doc.getRange().replace("{{date}}", getDateString(info.date().getDate()), new FindReplaceOptions());

            Table table = (Table) doc.getChild(NodeType.TABLE, 1, true);

            if (table.getRows().getCount() > 1) {
                table.getRows().get(1).remove();
            }

            Row templateRow = table.getFirstRow();

            int count = 0;
            for (var student : info.infoStudents().stream().filter(StudentResultInfo::isBest).toList()) {
                Row row = (Row) templateRow.deepClone(true);

                for (Cell cell : row.getCells()) {
                    cell.removeAllChildren();
                }
                count++;
                for (int i = 0; i < 5; i++) {
                    Cell cell = row.getCells().get(i);
                    Paragraph para = new Paragraph(doc);
                    Run run = new Run(doc);
                    run.getFont().setName("Arial");
                    run.getFont().setSize(10);

                    switch (i) {
                        case 0 -> run.setText(Integer.toString(count));
                        case 1 -> run.setText(student.getName());
                        case 2 -> run.setText(student.getThemeName());
                        case 3 -> run.setText(student.getSupervisorName());
                        case 4 -> run.setText("0");
                    }
                    para.appendChild(run);
                    cell.appendChild(para);
                }
                table.appendChild(row);
            }

            doc.getRange().replace("{{form_education}}", info.formEducation(), new FindReplaceOptions());
            doc.getRange().replace("{{qualification}}", info.qualification(), new FindReplaceOptions());

            var all = Double.parseDouble(String.valueOf(info.infoStudents().size()));
            doc.getRange().replace("{{all}}", String.valueOf(Math.round(all)), new FindReplaceOptions());
            doc.getRange().replace("{{allp}}", String.valueOf(100), new FindReplaceOptions());

            var red = info.infoStudents().stream().filter(StudentResultInfo::isRed).count();
            doc.getRange().replace("{{red}}", String.valueOf(red), new FindReplaceOptions());
            doc.getRange().replace("{{redp}}", String.valueOf(Math.round(red / all * 100)), new FindReplaceOptions());

            var otl = info.infoStudents().stream().filter(it -> it.getMark().equals(5)).count();
            doc.getRange().replace("{{otl}}", String.valueOf(otl), new FindReplaceOptions());
            doc.getRange().replace("{{otlp}}", String.valueOf(Math.round(otl / all * 100)), new FindReplaceOptions());

            var hor = info.infoStudents().stream().filter(it -> it.getMark().equals(4)).count();
            doc.getRange().replace("{{hor}}", String.valueOf(hor), new FindReplaceOptions());
            doc.getRange().replace("{{horp}}", String.valueOf(Math.round(hor / all * 100)), new FindReplaceOptions());

            var ud = info.infoStudents().stream().filter(it -> it.getMark().equals(3)).count();
            doc.getRange().replace("{{ud}}", String.valueOf(ud), new FindReplaceOptions());
            doc.getRange().replace("{{udp}}", String.valueOf(Math.round(ud / all * 100)), new FindReplaceOptions());

            var nud = info.infoStudents().stream().filter(it -> it.getMark().equals(2)).count();
            doc.getRange().replace("{{nud}}", String.valueOf(nud), new FindReplaceOptions());
            doc.getRange().replace("{{nudp}}", String.valueOf(Math.round(nud / all * 100)), new FindReplaceOptions());

            doc.save(outputStream, SaveFormat.DOCX);
            log.info("Документ успешно создан");
            return outputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Не удалось сгенерировать документ");
        }
    }

    private String getInverseName(String memberName) {
        if (memberName == null || memberName.trim().isEmpty()) {
            return memberName;
        }

        String[] parts = memberName.trim().split("\\s+");

        if (parts.length < 2) {
            return memberName;
        }

        if (parts.length == 2) {
            return parts[1] + " " + parts[0];
        } else {
            return parts[1] + " " + parts[2] + " " + parts[0];
        }
    }

    private String getMonth(int month) {
        return switch (month) {
            case 1 -> "января";
            case 2 -> "февраля";
            case 3 -> "марта";
            case 4 -> "апреля";
            case 5 -> "мая";
            case 6 -> "июня";
            case 7 -> "июля";
            case 8 -> "августа";
            case 9 -> "сентября";
            case 10 -> "октября";
            case 11 -> "ноября";
            case 12 -> "декабря";
            default -> throw new IllegalArgumentException("Некорректный номер месяца: " + month);
        };
    }
}

