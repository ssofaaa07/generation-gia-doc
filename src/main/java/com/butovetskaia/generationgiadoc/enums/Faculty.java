package com.butovetskaia.generationgiadoc.enums;

import lombok.Getter;

public enum Faculty {
    MILITARY_TRAINING_CENTER("Военный учебный центр"),
    FACULTY_OF_GEOGRAPHY_GEOECOLOGY_AND_TOURISM("Факультет географии, геоэкологии и туризма"),
    FACULTY_OF_GEOLOGY("Геологический факультет"),
    FACULTY_OF_JOURNALISM("Факультет журналистики"),
    FACULTY_OF_HISTORY("Исторический факультет"),
    FACULTY_OF_COMPUTER_SCIENCE("Факультет компьютерных наук"),
    FACULTY_OF_MATHEMATICS("Математический факультет"),
    FACULTY_OF_MEDICAL_AND_BIOLOGICAL_SCIENCES("Медико-биологический факультет"),
    FACULTY_OF_INTERNATIONAL_RELATIONS("Факультет международных отношений"),
    FACULTY_OF_APPLIED_MATHEMATICS_COMPUTER_SCIENCE_AND_MECHANICS("Факультет прикладной математики, информатики и механики"),
    FACULTY_OF_ROMANO_GERMANIC_PHILOLOGY("Факультет романо-германской филологии"),
    FACULTY_OF_PHARMACY("Фармацевтический факультет"),
    FACULTY_OF_PHYSICS("Физический факультет"),
    FACULTY_OF_PHILOLOGY("Филологический факультет"),
    FACULTY_OF_PHILOSOPHY_AND_PSYCHOLOGY("Факультет философии и психологии"),
    FACULTY_OF_CHEMISTRY("Химический факультет"),
    FACULTY_OF_ECONOMICS("Экономический факультет"),
    FACULTY_OF_LAW("Юридический факультет");

    @Getter
    private final String code;

    Faculty(String code) {
        this.code = code;
    }
}
