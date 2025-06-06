package com.butovetskaia.generationgiadoc.model;

import lombok.Builder;

import java.util.List;

@Builder
public record DocumentResultInfo(
        String direction,
        String faculty,
        String department,
        String qualification,
        String formEducation,
        Integer numberOfProtocol,
        DateInfo date,
        List<StudentResultInfo> infoStudents,
        CommissionMemberInfo secretaryName,
        CommissionMemberInfo chairpersonName,
        List<CommissionMemberInfo> commissionMembers,
        CommissionMemberInfo chairpersonAppellateName,
        List<CommissionMemberInfo> commissionAppellateMembers,
        boolean declineNames,
        int countStudentTheme,
        int countEnterpriseTheme,
        int countResearchTheme,
        int countPublicationRecommendation,
        int countImplementationRecommendation,
        int countImplemented
) {}
