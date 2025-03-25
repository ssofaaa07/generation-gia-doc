package com.butovetskaia.generationgiadoc.model;

import lombok.Builder;

import java.time.LocalDate;
import java.util.List;

@Builder
public record DocumentInfo (
        String direction,
        InfoCommissionMember secretaryName,
        InfoCommissionMember chairpersonName,
        List<InfoStudent> infoStudents,
        int annexNumber
) {}
