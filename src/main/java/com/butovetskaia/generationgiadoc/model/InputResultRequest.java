package com.butovetskaia.generationgiadoc.model;

import lombok.AllArgsConstructor;
import lombok.Getter;
import lombok.NoArgsConstructor;
import lombok.Setter;
import org.springframework.web.multipart.MultipartFile;

@AllArgsConstructor
@NoArgsConstructor
@Getter
@Setter
public class InputResultRequest {
    MultipartFile file;

    DateInfo schedule;

    boolean declineNames;

}
