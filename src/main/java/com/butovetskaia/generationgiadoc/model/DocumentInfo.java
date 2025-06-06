package com.butovetskaia.generationgiadoc.model;

import lombok.Builder;

import java.util.Collections;
import java.util.List;

@Builder
public record DocumentInfo (
        String direction,
        String faculty,
        String qualification,
        String formEducation,
        String department,
        List<DateInfo> schedules,
        CommissionMemberInfo secretaryName,
        CommissionMemberInfo chairpersonName,
        List<CommissionMemberInfo> commissionMembers,
        CommissionMemberInfo chairpersonAppellateName,
        List<CommissionMemberInfo> commissionAppellateMembers,
        List<StudentInfo> infoStudents,
        boolean declineNames
) {

    public static DocumentInfo of(DocumentResultInfo item) {
        return DocumentInfo.builder()
                .direction(item.direction())
                .faculty(item.faculty())
                .qualification(item.qualification())
                .formEducation(item.formEducation())
                .department(item.department())
                .secretaryName(item.secretaryName())
                .chairpersonName(item.chairpersonName())
                .commissionMembers(item.commissionMembers())
                .schedules(Collections.singletonList(item.date()))
                .declineNames(item.declineNames()).build();
    }
}
