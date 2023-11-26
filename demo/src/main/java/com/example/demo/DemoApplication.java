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

		ExcelParameter excelParameter = new ExcelParameter();
		GetExcelData getExcelData = new GetExcelData();

		excelParameter.setUrlOfExcel(getExcelData.getExcelUrl());
		excelParameter.setWorkSheet(getExcelData.getExcelWorkSheet());
		excelParameter.setColumnOfSum(getExcelData.getExcelColumnOfSum());
		excelParameter.setColumnOfElement(getExcelData.getExcelColumnOfElement());
		excelParameter.setColumnOfMass(getExcelData.getExcelColumnOfMass());
		excelParameter.setColumnOfRange(getExcelData.getExcelColumnOfRange());
		excelParameter.setElementOfAnaylis(getExcelData.getElementOfAnaylis());

		String urlOfExcel = excelParameter.getUrlOfExcel();
		Integer workSheet = excelParameter.getWorkSheet();
		Integer columnOfSum = excelParameter.getColumnOfSum();
		Integer columnOfElement = excelParameter.getColumnOfElement();
		Integer columnOfMass = excelParameter.getColumnOfMass();
		Integer columnOfRange = excelParameter.getColumnOfRange();
		ArrayList<String> elementOfAnaylis = getExcelData.getAnalysisEleList(excelParameter.getElementOfAnaylis());

		excelParameter
				.setNumOfParticle(getExcelData.getExcelDataSumOfNum(urlOfExcel, workSheet, columnOfSum, columnOfRange));

		// 取得每種材料的所有元素範圍
		ArrayList<Integer> numOfParticle = excelParameter.getNumOfParticle();
		System.out.println(numOfParticle);
		// 計算某規定上限物值的含量
		ArrayList<HashMap<String, Double>> elementMassData = getExcelData.getElementMassData(urlOfExcel, workSheet,
				columnOfElement, columnOfMass, columnOfRange);
		System.out.println(elementMassData.get(5));

		getExcelData.calculateWeightPercent(numOfParticle, elementOfAnaylis, elementMassData, urlOfExcel, workSheet,
				columnOfSum);

	}

}
