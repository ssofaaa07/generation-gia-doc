package com.butovetskaia.generationgiadoc.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class InputRequest {
    MultipartFile file;

    String faculty;

    String direction;

    String qualification;

    String formEducation;

    String department;

    List<DateInfo> schedules;

    CommissionMemberInfo secretary;

    CommissionMemberInfo chairperson;

    List<CommissionMemberInfo> commissionMembers;

    CommissionMemberInfo chairpersonAppellate;

    List<CommissionMemberInfo> commissionAppellateMembers;

    boolean declineNames;
}
