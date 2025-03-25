package com.butovetskaia.generationgiadoc.service.generation;

import com.aspose.cells.DateTime;
import com.aspose.words.*;
import com.butovetskaia.generationgiadoc.model.DocumentInfo;
import com.butovetskaia.generationgiadoc.model.InfoStudent;
import lombok.AllArgsConstructor;
import lombok.Setter;
import lombok.extern.slf4j.Slf4j;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;

@Slf4j
@AllArgsConstructor
public class AnnexProtocolGeneration extends DocumentGeneration {

    @Setter
    public int annexNumber;

    @Override
    public ByteArrayOutputStream generationDocument(DocumentInfo info, DateTime date) {
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            InputStream inputStream = getClass().getClassLoader().getResourceAsStream(DocumentGeneration.ANNEX_PROTOCOL_FILE_NAME + "-template.docx");
            Document template = new Document(inputStream);

            template.getRange().replace("{{count}}", String.valueOf(annexNumber), new FindReplaceOptions());
            template.getRange().replace("{{date}}", getDateString(date), new FindReplaceOptions());
            template.getRange().replace("{{chair_person}}", info.chairpersonName().getMemberNameShort(), new FindReplaceOptions());
            template.getRange().replace("{{secretary_person}}", info.secretaryName().getMemberNameShort(), new FindReplaceOptions());

            Document doc = new Document();

            for (InfoStudent student : info.infoStudents().stream().filter(is -> is.getDateExam().equals(date)).toList()) {

                Document tempDoc = template.deepClone();
                replacePlaceholders(tempDoc, student);
                doc.appendDocument(tempDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
//
//                appendDocument(tempDoc, doc);
//
//                if (info.infoStudents().indexOf(student) < info.infoStudents().size() - 1) {
//                    doc.getFirstSection().getBody().appendChild(new Paragraph(doc));
//                    doc.getFirstSection().getBody().getLastParagraph().appendChild(new Run(doc, "\f")); // Разрыв страницы
//                }
            }
            doc.save(outputStream, com.aspose.words.SaveFormat.DOCX);
            log.info("Документ успешно создан");
            return outputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("ggg");
        }
    }

    private void replacePlaceholders(Document doc, InfoStudent student) throws Exception {
        doc.getRange().replace("{{date}}", student.getDateExam().toString(), new FindReplaceOptions());
        doc.getRange().replace("{{student_name}}", student.getName(), new FindReplaceOptions());
        doc.getRange().replace("{{theme_name}}", student.getThemeName(), new FindReplaceOptions());
        doc.getRange().replace("{{supervisor_name}}", student.getSupervisorName(), new FindReplaceOptions());
    }

//    private void appendDocument(Document finalDoc, Document tempDoc) {
//        for (Section section : tempDoc.getSections()) {
//            NodeImporter importer = new NodeImporter(tempDoc, finalDoc, ImportFormatMode.KEEP_SOURCE_FORMATTING);
//            for (Node node : section.getBody()) {
//                Node importedNode = importer.importNode(node, true);
//                finalDoc.getFirstSection().getBody().appendChild(importedNode);
//            }
//        }
//    }
}
