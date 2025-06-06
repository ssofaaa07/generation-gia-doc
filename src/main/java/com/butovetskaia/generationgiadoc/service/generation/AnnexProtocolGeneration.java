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
import java.util.List;

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

            List<StudentInfo> filteredStudents = info.infoStudents().stream()
                    .filter(is -> is.getDateExam().equals(date))
                    .toList();

            for (int i = 0; i < filteredStudents.size(); i++) {
                StudentInfo student = filteredStudents.get(i);

                // Клонируем шаблон
                Document tempDoc = template.deepClone();
                replacePlaceholders(tempDoc, student, info.declineNames());

                // Получаем все параграфы и преобразуем в List вручную
                List<Paragraph> paragraphs = new ArrayList<>();
                for (Paragraph p : tempDoc.getFirstSection().getBody().getParagraphs()) {
                    paragraphs.add(p);
                }

                // Копируем оставшиеся параграфы в основной документ
                for (Paragraph sourceParagraph : paragraphs) {
                    Paragraph importedParagraph = (Paragraph) doc.importNode(sourceParagraph, true);
                    doc.getLastSection().getBody().appendChild(importedParagraph);
                }

                // Добавляем разрыв страницы (кроме последнего студента)
                if (i < filteredStudents.size() - 1) {
                    Paragraph pageBreak = new Paragraph(doc);
                    pageBreak.appendChild(new Run(doc, "\f")); // Форсированный разрыв страницы
                    doc.getLastSection().getBody().appendChild(pageBreak);
                }

                // Удаляем пустые параграфы в начале
                while (!paragraphs.isEmpty() && paragraphs.getFirst().getText().trim().isEmpty()) {
                    paragraphs.removeFirst();
                }

                // Удаляем пустые параграфы в конце
                while (!paragraphs.isEmpty() && paragraphs.getLast().getText().trim().isEmpty()) {
                    paragraphs.removeLast();
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
