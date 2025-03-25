package com.butovetskaia.generationgiadoc.model;

import lombok.AllArgsConstructor;
import lombok.Data;
import lombok.NoArgsConstructor;
import org.springframework.web.multipart.MultipartFile;

import java.util.List;

@AllArgsConstructor
@NoArgsConstructor
@Data
public class InputRequest {
    MultipartFile file;

    String faculty;

    String direction;

    String qualification;

    String formEducation;

    InfoCommissionMember secretary;

    InfoCommissionMember chairperson;

    List<InfoCommissionMember> commissionMembers;

}
