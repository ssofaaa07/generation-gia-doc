package com.butovetskaia.generationgiadoc.model;

import com.fasterxml.jackson.annotation.JsonFormat;
import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.AllArgsConstructor;
import lombok.Builder;
import lombok.Data;
import lombok.NoArgsConstructor;

import java.time.LocalDate;
import java.time.LocalTime;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class DateInfo {
    @JsonProperty("date")
    @JsonFormat(pattern = "yyyy-MM-dd")
    LocalDate date;

    @JsonProperty("start")
    @JsonFormat(pattern = "HH:mm")
    LocalTime start;

    @JsonProperty("end")
    @JsonFormat(pattern = "HH:mm")
    LocalTime end;

    @JsonProperty("orderDate")
    @JsonFormat(pattern = "yyyy-MM-dd")
    LocalDate orderDate;

    @JsonProperty("orderNumber")
    Integer orderNumber;
}
