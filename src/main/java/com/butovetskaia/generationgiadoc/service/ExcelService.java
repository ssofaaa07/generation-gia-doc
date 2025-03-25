package com.butovetskaia.generationgiadoc.service;

import com.aspose.cells.Cells;
import com.aspose.cells.DateTime;
import com.aspose.cells.Workbook;
import com.aspose.cells.Worksheet;
import com.butovetskaia.generationgiadoc.enums.Recommendation;
import com.butovetskaia.generationgiadoc.enums.Theme;
import com.butovetskaia.generationgiadoc.model.InfoStudent;
import lombok.extern.slf4j.Slf4j;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.util.ArrayList;
import java.util.List;

@Service
@Slf4j
public class ExcelService {

    public List<InfoStudent> getStudentInfo(MultipartFile file) {
        try {
            Workbook workbook = new Workbook(file.getInputStream());

            List<InfoStudent> infoStudents = new ArrayList<>();

            for (int i = 0; i < workbook.getWorksheets().getCount(); i++) {
                Worksheet worksheet = workbook.getWorksheets().get(i);
                Cells cells = worksheet.getCells();

                infoStudents.addAll(readData(cells));
            }

            return infoStudents;
        } catch (Exception e) {
            e.printStackTrace();
        }
        return null;
    }

    private List<InfoStudent> readData(Cells cells) {
        List<InfoStudent> infoStudents = new ArrayList<>();

        DateTime date = cells.get(0, 2).getDateTimeValue();
        for (int row = 4; row < cells.getMaxDataRow() + 1; row++) {

            var themeType = getTheme(cells, row);
            var recommendations = getRecommendations(cells, row);

            InfoStudent info = InfoStudent.builder()
                    .dateExam(date)
                    .name(cells.get(row, 1).getStringValue())
                    .themeName(cells.get(row, 2).getStringValue())
                    .supervisorName(cells.get(row, 3).getStringValue())
                    .supervisorMark(cells.get(row, 4).getStringValue())
                    .themeType(themeType)
                    .recommendations(recommendations).build();
            infoStudents.add(info);
        }
        return infoStudents;
    }

    private Theme getTheme(Cells cells, int row) {
        if (cells.get(row, 6).getStringValue() != null && cells.get(row, 6).getStringValue().length() > 0) {
            return Theme.ENTERPRISE_THEME;
        } else if (cells.get(row, 7).getStringValue() != null && cells.get(row, 7).getStringValue().length() > 0) {
            return Theme.RESEARCH_THEME;
        }
        return Theme.STUDENT_THEME;
    }

    private List<Recommendation> getRecommendations(Cells cells, int row) {
        List<Recommendation> recommendations = new ArrayList<>();
        if (cells.get(row, 8).getStringValue() != null && cells.get(row, 8).getStringValue().length() > 0) {
            recommendations.add(Recommendation.PUBLICATION_RECOMMENDATION);
        }
        if (cells.get(row, 9).getStringValue() != null && cells.get(row, 9).getStringValue().length() > 0) {
            recommendations.add(Recommendation.IMPLEMENTATION_RECOMMENDATION);
        }
        if (cells.get(row, 10).getStringValue() != null && cells.get(row, 10).getStringValue().length() > 0) {
            recommendations.add(Recommendation.IMPLEMENTED);
        }
        return recommendations;
    }
}
