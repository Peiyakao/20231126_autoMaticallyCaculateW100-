package com.example.demo;

import java.util.ArrayList;
import java.util.HashMap;

import org.springframework.boot.SpringApplication;
import org.springframework.boot.autoconfigure.SpringBootApplication;

@SpringBootApplication
public class DemoApplication {

	public static void main(String[] args) {
		SpringApplication.run(DemoApplication.class, args);
		System.out.println("Start Analysis");
		GetExcelData getExcelData = new GetExcelData();
		getExcelData.getExcelUrl();

	}

}
