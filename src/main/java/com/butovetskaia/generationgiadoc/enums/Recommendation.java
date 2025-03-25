package com.butovetskaia.generationgiadoc.enums;

public enum Recommendation {
    PUBLICATION_RECOMMENDATION("ВКР рекомендована к опубликованию"),
    IMPLEMENTATION_RECOMMENDATION("ВКР рекомендована к внедрению"),
    IMPLEMENTED("ВКР внедрена");

    private String recommendation;

    Recommendation(String recommendation) {
        this.recommendation = recommendation;
    }
}
