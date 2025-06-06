package com.butovetskaia.generationgiadoc.model;

import lombok.Builder;
import lombok.Data;

@Data
@Builder
public class StudentResultInfo {
    String name;
    String themeName;
    String supervisorName;
    Integer mark;
    boolean isRed;
    boolean isBest;
}
