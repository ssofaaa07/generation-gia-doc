package com.butovetskaia.generationgiadoc.model;

import com.butovetskaia.generationgiadoc.enums.Recommendation;
import com.butovetskaia.generationgiadoc.enums.Theme;
import lombok.Builder;
import lombok.Data;

import java.util.List;

@Data
@Builder
public class StudentInfo {
    DateInfo dateExam;
    String name;
    String themeName;
    String supervisorName;
    String supervisorMark;
    Theme themeType;
    List<Recommendation> recommendations;
    int markOne;
    int markTwo;
}
