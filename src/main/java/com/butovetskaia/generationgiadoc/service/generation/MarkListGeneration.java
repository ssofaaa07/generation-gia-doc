package com.butovetskaia.generationgiadoc.service.generation;

import com.aspose.words.*;
import com.butovetskaia.generationgiadoc.enums.Recommendation;
import com.butovetskaia.generationgiadoc.model.DateInfo;
import com.butovetskaia.generationgiadoc.model.DocumentInfo;
import com.butovetskaia.generationgiadoc.model.StudentInfo;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

@Slf4j
public class MarkListGeneration extends DocumentGeneration {
    @Override
    public ByteArrayOutputStream generationDocument(DocumentInfo info, DateInfo date) {
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            InputStream inputStream = getClass().getClassLoader().getResourceAsStream(DocumentGeneration.MARK_LIST_FILE_NAME + "-template.docx");
            Document doc = new Document(inputStream);

            doc.getRange().replace("{{direction}}", info.direction(), new FindReplaceOptions());

            Table table = (Table) doc.getChild(NodeType.TABLE, 0, true);

            if (table.getRows().getCount() > 2) {
                table.getRows().get(2).remove();
            }
            int count = 0;

            for (StudentInfo student : info.infoStudents().stream().filter(is -> is.getDateExam().getDate().equals(date.getDate())).toList()) {
                Row row = new Row(doc);
                count++;
                for (int i = 1; i < 16; i++) {
                    Cell cell = new Cell(doc);
                    Paragraph para = new Paragraph(doc);
                    switch (i) {
                        case 1 -> para.appendChild(new Run(doc, Integer.toString(count)));
                        case 2 -> para.appendChild(new Run(doc, student.getName()));
                        case 3 -> para.appendChild(new Run(doc, student.getThemeName()));
                        case 4 -> para.appendChild(new Run(doc, student.getSupervisorName()));
                        case 5 -> para.appendChild(new Run(doc, student.getSupervisorMark()));
                        case 6 -> para.appendChild(new Run(doc, student.getThemeType().isStudentTheme() ? "+" : ""));
                        case 7 -> para.appendChild(new Run(doc, student.getThemeType().isEnterpriseTheme() ? "+" : ""));
                        case 8 -> para.appendChild(new Run(doc, student.getThemeType().isResearchTheme() ? "+" : ""));
                        case 9 ->
                                para.appendChild(new Run(doc, student.getRecommendations().contains(Recommendation.PUBLICATION_RECOMMENDATION) ? "+" : ""));
                        case 10 ->
                                para.appendChild(new Run(doc, student.getRecommendations().contains(Recommendation.IMPLEMENTATION_RECOMMENDATION) ? "+" : ""));
                        case 11 ->
                                para.appendChild(new Run(doc, student.getRecommendations().contains(Recommendation.IMPLEMENTED) ? "+" : ""));
                        case 12 -> para.appendChild(new Run(doc, student.getMarkOne()));
                        case 13 -> para.appendChild(new Run(doc, student.getMarkTwo()));
                    }
                    cell.appendChild(para);
                    row.appendChild(cell);
                }
                table.appendChild(row);
            }

            doc.save(outputStream, com.aspose.words.SaveFormat.DOCX);
            log.info("Документ успешно создан");

            return outputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Не удалось сгенерировать документ");
        }
    }
}
