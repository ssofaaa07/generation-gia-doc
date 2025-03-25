package com.butovetskaia.generationgiadoc;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;
import org.springframework.context.annotation.Bean;
import org.springframework.web.filter.HiddenHttpMethodFilter;

@SpringBootApplication
public class GenerationGiaDocApplication {

	public static void main(String[] args) {
		SpringApplication.run(GenerationGiaDocApplication.class, args);
	}
}
