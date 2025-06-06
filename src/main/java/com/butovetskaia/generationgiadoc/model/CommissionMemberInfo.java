package com.butovetskaia.generationgiadoc.model;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class CommissionMemberInfo {
    @JsonProperty("memberName")
    String memberName;

    @JsonProperty("memberPost")
    String memberPost;

    @JsonProperty("orderNumber")
    Integer orderNumber;

    @JsonProperty("orderDate")
    @JsonFormat(pattern = "yyyy-MM-dd")
    LocalDate orderDate;
}
