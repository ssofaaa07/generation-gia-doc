package com.butovetskaia.generationgiadoc.service.generation;

import com.aspose.cells.DateTime;
import com.butovetskaia.generationgiadoc.model.DocumentInfo;
import com.butovetskaia.generationgiadoc.model.InfoStudent;
import lombok.RequiredArgsConstructor;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

@RequiredArgsConstructor
public abstract class DocumentGeneration {

    protected static final String MARK_LIST_FILE_NAME = "Оценочный лист";
    protected static final String ANNEX_PROTOCOL_FILE_NAME = "Приложение к протоколу ГЭК о защите";

    public static ByteArrayOutputStream generation(DocumentInfo info) throws IOException {
        MarkListGeneration markListGeneration = new MarkListGeneration();
        AnnexProtocolGeneration annexProtocolGeneration = new AnnexProtocolGeneration(1);

        ByteArrayOutputStream zipOutputStream = new ByteArrayOutputStream();
        try (ZipOutputStream zipOut = new ZipOutputStream(zipOutputStream)) {

            var dateExamList = info.infoStudents().stream().map(InfoStudent::getDateExam).distinct().sorted().toList();
            for (DateTime date : dateExamList) {
                ByteArrayOutputStream markListOutputStream = markListGeneration.generationDocument(info, date);
                zipOut.putNextEntry(new ZipEntry(MARK_LIST_FILE_NAME + " " + getDateString(date) + ".docx"));
                zipOut.write(markListOutputStream.toByteArray());
                zipOut.closeEntry();

                ByteArrayOutputStream annexProtocolOutputStream = annexProtocolGeneration.generationDocument(info, date);
                zipOut.putNextEntry(new ZipEntry(ANNEX_PROTOCOL_FILE_NAME + " " + getDateString(date) + ".docx"));
                zipOut.write(annexProtocolOutputStream.toByteArray());
                zipOut.closeEntry();
                annexProtocolGeneration.setAnnexNumber(annexProtocolGeneration.annexNumber + 1);
            }
            return zipOutputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("dgfsjfg");
        }
    }

    public abstract ByteArrayOutputStream generationDocument(DocumentInfo info, DateTime date);

    public static String getDateString(DateTime date) {
        var day = String.valueOf(date.getDay());
        var month = String.valueOf(date.getMonth());
        var year = String.valueOf(date.getYear());

        if (day.length() < 2) {
            day = "0" + day;
        }
        if (month.length() < 2) {
            month = "0" + month;
        }
        return day + "." + month + "." + year;
    }
}
