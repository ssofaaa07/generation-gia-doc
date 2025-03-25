package com.butovetskaia.generationgiadoc.enums;

public enum Theme {
    STUDENT_THEME("ВКР выполнена по теме, предложенной студентом"),
    ENTERPRISE_THEME("ВКР выполнена по заявке предприятия/организации"),
    RESEARCH_THEME("ВКР относится к области фундаментальных и поисковых научных исследований");

    private String theme;

    Theme(String theme) {
        this.theme = theme;
    }

    public boolean isStudentTheme() {
        return this == STUDENT_THEME;
    }

    public boolean isEnterpriseTheme() {
        return this == ENTERPRISE_THEME;
    }

    public boolean isResearchTheme() {
        return this == RESEARCH_THEME;
    }
}
