package com.butovetskaia.generationgiadoc.service.generation;

import com.aspose.words.ControlChar;
import com.aspose.words.Document;
import com.aspose.words.FindReplaceOptions;
import com.aspose.words.SaveFormat;
import com.butovetskaia.generationgiadoc.model.DateInfo;
import com.butovetskaia.generationgiadoc.model.DocumentResultInfo;
import com.butovetskaia.generationgiadoc.model.StudentResultInfo;
import lombok.extern.slf4j.Slf4j;
import ru.morpher.ws3.AccessDeniedException;
import ru.morpher.ws3.ArgumentEmptyException;
import ru.morpher.ws3.ClientBuilder;
import ru.morpher.ws3.russian.ArgumentNotRussianException;
import ru.morpher.ws3.russian.InvalidFlagsException;
import ru.morpher.ws3.russian.NumeralsDeclensionNotSupportedException;
import ru.morpher.ws3.russian.RussianClient;

import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Comparator;
import java.util.List;
import java.util.Objects;
import java.util.stream.Collectors;

import static com.butovetskaia.generationgiadoc.service.generation.DocumentGeneration.getDateString;
import static com.butovetskaia.generationgiadoc.service.generation.DocumentGeneration.getShortName;

@Slf4j
public class AssignmentOfQualificationGeneration {

    public ByteArrayOutputStream generationDocument(DocumentResultInfo info, DateInfo date) {
        try {
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

            InputStream inputStream = getClass().getClassLoader().getResourceAsStream(DocumentGeneration.ASSIGNMENT_OF_QUALIFICATION_FILE_NAME + "-template.doc");
            Document doc = new Document(inputStream);

            RussianClient morpher = info.declineNames() ? new ClientBuilder().useToken("6b9f45d6-5443-4c13-b4ea-c0e8d5248b72").build().russian() : null;

            doc.getRange().replace("{{count}}", String.valueOf(info.numberOfProtocol()), new FindReplaceOptions());
            doc.getRange().replace("{{date}}", getDateString(date.getDate()), new FindReplaceOptions());
            doc.getRange().replace("{{direction}}", info.direction(), new FindReplaceOptions());
            doc.getRange().replace("{{year}}", String.valueOf(info.date().getDate().getYear()), new FindReplaceOptions());
            doc.getRange().replace("{{qualification}}", info.qualification() + "а", new FindReplaceOptions());
            doc.getRange().replace("{{form_education}}", info.formEducation().toLowerCase(), new FindReplaceOptions());
            doc.getRange().replace("{{chairperson}}", getShortName(info.chairpersonName().getMemberName()), new FindReplaceOptions());
            doc.getRange().replace("{{secretary}}", getShortName(info.secretaryName().getMemberName()), new FindReplaceOptions());

            var faculty = info.declineNames() ? morpher.declension(info.faculty().toLowerCase()).genitive : info.faculty().replaceFirst(" ", "а ").toLowerCase();
            doc.getRange().replace("{{faculty}}", faculty, new FindReplaceOptions());

            var countStudents = info.infoStudents().size();
            doc.getRange().replace("{{count_students}}", String.valueOf(countStudents), new FindReplaceOptions());

            var redStudents = info.infoStudents().stream()
                    .filter(StudentResultInfo::isRed)
                    .sorted(Comparator.comparing(StudentResultInfo::getName))
                    .toList();

            var regularStudents = info.infoStudents().stream()
                    .filter(it -> !it.isRed())
                    .sorted(Comparator.comparing(StudentResultInfo::getName))
                    .toList();

            // Формируем полный список студентов с правильной пунктуацией
            StringBuilder studentsBuilder = new StringBuilder();

            if (!redStudents.isEmpty()) {
                studentsBuilder.append("дипломы с отличием" + ControlChar.LINE_BREAK_CHAR)
                        .append(formatStudentList(redStudents, info.declineNames(), morpher, false))
                        .append(ControlChar.LINE_BREAK_CHAR + "" + ControlChar.LINE_BREAK_CHAR + "дипломы" + ControlChar.LINE_BREAK_CHAR);
            } else {
                studentsBuilder.append("дипломы" + ControlChar.LINE_BREAK_CHAR);
            }

            studentsBuilder.append(formatStudentList(regularStudents, info.declineNames(), morpher, true));

            // Заменяем оба плейсхолдера одним сформированным текстом
            doc.getRange().replace("{{student_name}}", studentsBuilder.toString(), new FindReplaceOptions());

            doc.save(outputStream, SaveFormat.DOCX);
            log.info("Документ успешно создан");
            return outputStream;
        } catch (Exception e) {
            e.printStackTrace();
            throw new RuntimeException("Не удалось сгенерировать документ");
        }
    }

    private String formatStudentList(List<StudentResultInfo> students, boolean declineNames,
                                     RussianClient morpher, boolean addDot) {
        if (students.isEmpty()) {
            return "";
        }

        String list = students.stream()
                .map(student -> {
                    try {
                        return declineNames
                                ? Objects.requireNonNull(morpher).declension(student.getName()).dative
                                : student.getName().trim();
                    } catch (IOException | AccessDeniedException | ArgumentEmptyException |
                             InvalidFlagsException | ArgumentNotRussianException |
                             NumeralsDeclensionNotSupportedException e) {
                        throw new RuntimeException(e);
                    }
                })
                .collect(Collectors.joining(", " + ControlChar.LINE_BREAK_CHAR));

        return list + (addDot ? "." : ",");
    }
}

