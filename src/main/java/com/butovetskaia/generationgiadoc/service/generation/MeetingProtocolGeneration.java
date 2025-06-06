package com.butovetskaia.generationgiadoc.service.generation;

import com.aspose.words.*;
import com.butovetskaia.generationgiadoc.model.CommissionMemberInfo;
import com.butovetskaia.generationgiadoc.model.DateInfo;
import com.butovetskaia.generationgiadoc.model.DocumentInfo;
import lombok.AllArgsConstructor;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

@Slf4j
@AllArgsConstructor
public class MeetingProtocolGeneration extends DocumentGeneration {

    @Setter
    public int protocolNumber;
    @Setter
    public boolean isAssignmentOfQualification;

    @Override
    public ByteArrayOutputStream generationDocument(DocumentInfo info, DateInfo date) {
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            InputStream inputStream = getClass().getClassLoader().getResourceAsStream(DocumentGeneration.MEETING_PROTOCOL_FILE_NAME + "-template.doc");
            Document doc = new Document(inputStream);

            doc.getRange().replace("{{count}}", String.valueOf(protocolNumber), new FindReplaceOptions());
            doc.getRange().replace("{{date}}", getDateString(date.getDate()), new FindReplaceOptions());
            doc.getRange().replace("{{start}}", String.valueOf(date.getStart().getHour()), new FindReplaceOptions());
            doc.getRange().replace("{{end}}", String.valueOf(date.getEnd().getHour()), new FindReplaceOptions());
            doc.getRange().replace("{{direction}}", info.direction(), new FindReplaceOptions());
            doc.getRange().replace("{{chairperson_name}}", info.chairpersonName().getMemberName(), new FindReplaceOptions());
            doc.getRange().replace("{{chairperson_post}}", info.chairpersonName().getMemberPost(), new FindReplaceOptions());
            doc.getRange().replace("{{secretary_name}}", info.secretaryName().getMemberName(), new FindReplaceOptions());
            doc.getRange().replace("{{secretary_post}}", info.secretaryName().getMemberPost(), new FindReplaceOptions());
            doc.getRange().replace("{{chairperson}}", getShortName(info.chairpersonName().getMemberName()), new FindReplaceOptions());
            doc.getRange().replace("{{secretary}}", getShortName(info.secretaryName().getMemberName()), new FindReplaceOptions());

            var title = "";
            var agendas = "";
            var directionWithoutCode = info.direction().substring(info.direction().indexOf(" ") + 1);
            if (isAssignmentOfQualification) {
                title = "о присвоении квалификации выпускникам";
                agendas = "присвоение квалификации выпускникам направления " + info.direction();
            } else {
                title = "по программе";
                agendas = "защита выпускных квалификационных работ " + info.qualification().toLowerCase() + "ов, направление " + directionWithoutCode;
            }
            doc.getRange().replace("{{agendas}}", agendas, new FindReplaceOptions());
            doc.getRange().replace("{{title}}", title, new FindReplaceOptions());

            if (isAssignmentOfQualification) {
                doc.getRange().replace("{{count_annex}}", "№ 1", new FindReplaceOptions());
            } else {
                var countAnnex = info.infoStudents().stream().filter(is -> is.getDateExam().equals(date)).count();
                doc.getRange().replace("{{count_annex}}", "№№ 1-" + countAnnex, new FindReplaceOptions());
            }

            Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);
            Row templateRow = table.getRows().get(0);

            if (!info.commissionMembers().isEmpty()) {
                table.getRows().get(0).remove();
            }

            for (CommissionMemberInfo member : info.commissionMembers()) {
                Row newRow = (Row) templateRow.deepClone(true);

                for (Cell cell : newRow.getCells()) {
                    for (Paragraph para : cell.getParagraphs()) {
                        String text = para.getText();
                        var runs = para.getRuns();
                        if (text.contains("commission_member_name")) {
                            runs.get(0).setText(text.replace("commission_member_name", member.getMemberName()));
                        } else if (text.contains("commission_member_post")) {
                            runs.get(0).setText(text.replace("commission_member_post", member.getMemberPost()));
                        }
                        var i = 0;
                        for (Run run : runs) {
                            if (i != 0) run.remove();
                            i++;
                        }
                    }
                }
                table.appendChild(newRow);
            }

            var lenRuns = table.getLastRow().getLastCell().getLastParagraph().getRuns().getCount();
            var text = table.getLastRow().getLastCell().getLastParagraph().getRuns().get(lenRuns - 1).getText();
            log.info(text);
            table.getLastRow().getLastCell().getLastParagraph().getRuns().get(lenRuns - 1).setText(text.replace(";", "."));

            doc.save(outputStream, com.aspose.words.SaveFormat.DOCX);
            log.info("Документ успешно создан");
            return outputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Не удалось сгенерировать документ");
        }
    }
}

