package com.kanseiu.office;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

/**
 * User: kanseiu
 * Date: 2024/4/18
 * Project: OfficeConverter
 * Package: com.kanseiu.office
 */
@SpringBootApplication(scanBasePackages = "com.kanseiu.office")
public class OfficeConverterApplication {

    public static void main(String[] args) {
        SpringApplication.run(OfficeConverterApplication.class, args);
    }

}
