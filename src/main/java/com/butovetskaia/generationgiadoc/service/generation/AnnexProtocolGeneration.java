package com.butovetskaia.generationgiadoc.service.generation;

import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.Paragraph;
import com.aspose.words.Run;
import com.butovetskaia.generationgiadoc.model.DateInfo;
import com.butovetskaia.generationgiadoc.model.DocumentInfo;
import com.butovetskaia.generationgiadoc.model.StudentInfo;
import lombok.AllArgsConstructor;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;
import ru.morpher.ws3.ClientBuilder;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;
import java.util.stream.Collectors;


@Slf4j
@AllArgsConstructor
public class AnnexProtocolGeneration extends DocumentGeneration {

    @Setter
    public int annexNumber;

    @Override
    public ByteArrayOutputStream generationDocument(DocumentInfo info, DateInfo date) {
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            InputStream inputStream = getClass().getClassLoader().getResourceAsStream(DocumentGeneration.ANNEX_PROTOCOL_FILE_NAME + "-template.docx");
            Document template = new Document(inputStream);

            template.getRange().replace("{{count}}", String.valueOf(annexNumber), new FindReplaceOptions());
            template.getRange().replace("{{date}}", getDateString(date.getDate()), new FindReplaceOptions());
            template.getRange().replace("{{chairperson}}", getShortName(info.chairpersonName().getMemberName()), new FindReplaceOptions());
            template.getRange().replace("{{secretary}}", getShortName(info.secretaryName().getMemberName()), new FindReplaceOptions());

            Document doc = new Document();

            // Фильтруем студентов по дате
            List<StudentInfo> filteredStudents = info.infoStudents().stream()
                    .filter(is -> is.getDateExam().equals(date))
                    .toList();

            for (int i = 0; i < filteredStudents.size(); i++) {
                StudentInfo student = filteredStudents.get(i);

                // Клонируем и заполняем шаблон для текущего студента
                Document studentDoc = template.deepClone();
                replacePlaceholders(studentDoc, student, info.declineNames());

                // Получаем все параграфы и очищаем от пустых в начале/конце
                List<Paragraph> paragraphs = Arrays.stream(studentDoc.getFirstSection().getBody().getParagraphs().toArray())
                        .map(p -> (Paragraph)p)
                        .collect(Collectors.toList());

                // Удаляем пустые параграфы в начале
                while (!paragraphs.isEmpty() && paragraphs.get(0).getText().trim().isEmpty()) {
                    paragraphs.remove(0);
                }

                // Удаляем пустые параграфы в конце
                while (!paragraphs.isEmpty() && paragraphs.get(paragraphs.size()-1).getText().trim().isEmpty()) {
                    paragraphs.remove(paragraphs.size()-1);
                }

                // Для всех кроме первого студента добавляем разрыв страницы
                if (i > 0) {
                    Paragraph first = (Paragraph)doc.importNode(paragraphs.get(0), true);
                    first.getParagraphFormat().setPageBreakBefore(true);
                    doc.getLastSection().getBody().appendChild(first);

                    // Добавляем остальные параграфы
                    for (int j = 1; j < paragraphs.size(); j++) {
                        doc.getLastSection().getBody().appendChild(
                                doc.importNode(paragraphs.get(j), true));
                    }
                } else {
                    // Для первого студента просто добавляем все параграфы
                    for (Paragraph paragraph : paragraphs) {
                        doc.getLastSection().getBody().appendChild(
                                doc.importNode(paragraph, true));
                    }
                    doc.getFirstSection().getBody().getFirstChild().remove();
                }
            }
            doc.save(outputStream, com.aspose.words.SaveFormat.DOCX);
            log.info("Документ успешно создан");
            return outputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Не удалось сгенерировать документ");
        }
    }

    private void replacePlaceholders(Document doc, StudentInfo student, boolean declineNames) throws Exception {
        var studentNameIns = "";
        var supervizorName = "";
        if (declineNames) {
            var morpher = new ClientBuilder().useToken("6b9f45d6-5443-4c13-b4ea-c0e8d5248b72").build().russian();
            studentNameIns = morpher.declension(student.getName()).genitive;
            supervizorName = morpher.declension(getShortName(student.getSupervisorName())).genitive;
        } else {
            studentNameIns = student.getName();
            supervizorName = getShortName(student.getSupervisorName());
        }

        doc.getRange().replace("{{date}}", student.getDateExam().toString(), new FindReplaceOptions());
        doc.getRange().replace("{{student_name}}", student.getName(), new FindReplaceOptions());
        doc.getRange().replace("{{student_name_ins}}", studentNameIns, new FindReplaceOptions());
        doc.getRange().replace("{{theme_name}}", student.getThemeName(), new FindReplaceOptions());
        doc.getRange().replace("{{supervisor_name}}", supervizorName, new FindReplaceOptions());
    }
}
