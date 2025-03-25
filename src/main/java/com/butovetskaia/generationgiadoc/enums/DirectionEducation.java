package com.butovetskaia.generationgiadoc.enums;

import java.util.List;

public enum DirectionEducation {
    INFORMATION_SYSTEM_AND_TECHNOLOGIES("09.03.02","Информационные системы и технологии"),
    SOFTWARE_ENGINEERING("09.03.04","Программная инженерия"),
    MATHEMATICS_AND_COMPUTER_SCIENCE("02.03.01","Математика и компьютерные науки"),
    APPLIED_COMPUTER_SCIENCE("09.03.03","Прикладная информатика"),
    INFORMATION_SECURITY("10.03.01","Информационная безопасность");

    final String code;
    final String name;

    DirectionEducation(String code, String name) {
        this.code = code;
        this.name = name;
    }

    public static List<String> getStringValues() {
        List<DirectionEducation> directionEducations = List.of(DirectionEducation.values());
        return directionEducations.stream().map(DirectionEducation::getStringValue).toList();
    }

    public String getStringValue() {
        return code + " " + name;
    }
}
