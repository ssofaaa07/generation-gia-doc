package com.butovetskaia.generationgiadoc.model;

import com.fasterxml.jackson.annotation.JsonProperty;
import lombok.*;

@Data
@Builder
@AllArgsConstructor
@NoArgsConstructor
public class InfoCommissionMember {
//    @NonNull
    @JsonProperty("memberName")
    String memberName;
//    @NonNull
    @JsonProperty("memberPost")
    String memberPost;

    public String getMemberNameShort() {
        String[] parts = memberName.trim().split("\\s+");

        if (parts.length < 2) {
            throw new IllegalArgumentException("Full name must contain at least two parts: last name and first name");
        }

        String lastName = parts[0];
        String firstNameInitial = parts[1].charAt(0) + ".";

        if (parts.length >= 3) {
            String patronymicInitial = parts[2].charAt(0) + ".";
            return lastName + " " + firstNameInitial + patronymicInitial;
        } else {
            return lastName + " " + firstNameInitial;
        }
    }
}
