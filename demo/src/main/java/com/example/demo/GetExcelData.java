package com.example.demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.HashMap;

import org.apache.poi.ss.usermodel.*;
import java.util.Scanner;

public class GetExcelData {
    public String getExcelUrl() {
        String urlOfExcel;
        Scanner inputUrl = new Scanner(System.in);
        System.out.println("請輸入分析檔案路徑:");
        urlOfExcel = inputUrl.nextLine();
        // inputUrl.close();
        return urlOfExcel;
    }

    public Integer getExcelWorkSheet() {
        Integer workSheet;
        Scanner inputworkSheet = new Scanner(System.in);
        System.out.println("請輸入分析檔案工作表，如果是第一個工作表請填0:");
        workSheet = Integer.parseInt(inputworkSheet.nextLine());
        // inputworkSheet.close();
        return workSheet;

    }

    public Integer getExcelColumnOfSum() {
        Integer columnOfSum;
        Scanner inputSum = new Scanner(System.in);
        System.out.println("請輸入總和欄位，如果是A欄填寫0、B欄填寫1:");
        columnOfSum = Integer.parseInt(inputSum.nextLine());
        // inputSum.close();
        return columnOfSum;

    }

    public Integer getExcelColumnOfElement() {
        Integer columnOfElement;
        Scanner inputElement = new Scanner(System.in);
        System.out.println("請輸入存放元素欄位，如果是A欄填寫0、B欄填寫1:");
        columnOfElement = Integer.parseInt(inputElement.nextLine());
        // inputElement.close();
        return columnOfElement;
    }

    public Integer getExcelColumnOfMass() {
        Integer columnOfMass;
        Scanner inputMass = new Scanner(System.in);
        System.out.println("請輸入重量欄位，如果是A欄填寫0、B欄填寫1:");
        columnOfMass = Integer.parseInt(inputMass.nextLine());
        // inputMass.close();
        return columnOfMass;
    }

    public Integer getExcelColumnOfRange() {
        Integer columnOfRange;
        Scanner inputRange = new Scanner(System.in);
        System.out.println("請輸入資料範圍，如果數值最後一行為16，請填16:");
        columnOfRange = Integer.parseInt(inputRange.nextLine());
        // inputMass.close();
        return columnOfRange;
    }

    public String getElementOfAnaylis() {
        String elementOfAnaylis;
        Scanner inputelement = new Scanner(System.in);
        System.out.println("請輸入計算的對象名稱，並用'，'區隔，如:Ag,K,Ca:");
        elementOfAnaylis = inputelement.nextLine();
        inputelement.close();
        return elementOfAnaylis;
    }

    public ArrayList<Integer> getExcelDataSumOfNum(String urlOfExcel, Integer workSheet, Integer columnOfSum,
            Integer columnOfRange) {
        // 載入EXCEL資料
        ArrayList<Integer> numOfParticle = new ArrayList();
        try (InputStream excelFile = new FileInputStream(urlOfExcel)) {
            Workbook workbook = WorkbookFactory.create(excelFile);
            Sheet sheet = workbook.getSheetAt(workSheet);
            System.out.println("讀取最小物質數量");
            Integer temp = 1;
            for (int rowNum = 2; rowNum < columnOfRange; rowNum++) {

                Row sumRow = sheet.getRow(rowNum);
                Cell cell = sumRow.getCell(columnOfSum);
                CellType cellType = cell.getCellType();
                System.out.println(cellType);
                if (cellType == CellType.BLANK) {
                    temp++;
                    System.out.println(temp);

                } else if (cellType == CellType.NUMERIC) {
                    numOfParticle.add(temp);
                    temp = 1;
                }
                if (rowNum == columnOfRange - 1) {
                    System.out.println(cell.getStringCellValue());
                    numOfParticle.add(temp);
                    temp = 1;
                    workbook.close();
                    return numOfParticle;
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        return numOfParticle;
    }

    public ArrayList<String> getAnalysisEleList(String element) {
        String[] strArray = element.split(",");
        ArrayList<String> analysisElement = new ArrayList<>(Arrays.asList(strArray));
        return analysisElement;

    }

    public ArrayList<HashMap<String, Double>> getElementMassData(String urlOfExcel, Integer workSheet,
            Integer columnOfElement,
            Integer columnOfMass,
            Integer columnOfRange) {
        ArrayList<HashMap<String, Double>> elementMassData = new ArrayList<>();
        try (InputStream excelFile = new FileInputStream(urlOfExcel)) {
            Workbook workbook = WorkbookFactory.create(excelFile);
            Sheet sheet = workbook.getSheetAt(workSheet);
            System.out.println("讀取物質名稱及數值");
            for (int rowNum = 1; rowNum < columnOfRange; rowNum++) {
                Row elementAndMassRow = sheet.getRow(rowNum);
                Cell element = elementAndMassRow.getCell(columnOfElement);
                Cell mass = elementAndMassRow.getCell(columnOfMass);
                HashMap<String, Double> tempElement = new HashMap<>();
                tempElement.put(element.getStringCellValue(), mass.getNumericCellValue());
                System.out.println(tempElement.values());
                elementMassData.add(tempElement);

            }
        } catch (Exception e) {
            e.printStackTrace();
        }
        return elementMassData;
    }

    public void calculateWeightPercent(
            ArrayList<Integer> numOfParticle,
            ArrayList<String> elementOfAnaylis,
            ArrayList<HashMap<String, Double>> elementMassData,
            String urlOfExcel,
            Integer workSheet,
            Integer columnOfSum) {
        Integer tempSumofRow = 1;
        Integer writeRow = 0;
        try (InputStream excelFile = new FileInputStream(urlOfExcel)) {
            Workbook workbook = WorkbookFactory.create(excelFile);
            Sheet sheet = workbook.getSheetAt(workSheet);
            Sheet newSheet = workbook.createSheet();
            CellStyle percentageStyle = workbook.createCellStyle();
            percentageStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
            for (Integer num : numOfParticle) {
                Row sumRow = sheet.getRow(tempSumofRow);
                Cell cell = sumRow.getCell(columnOfSum);
                System.out.println("RowIndex:" + cell.getRowIndex());
                System.out.println(cell.getCellType());
                System.out.println("currentSum:" + cell.getNumericCellValue());
                Double currentSum = cell.getNumericCellValue();
                Integer tempWriteRow = writeRow;
                for (String analysisElement : elementOfAnaylis) {
                    Double accumulation = 0.00;
                    System.out.println("目前分析元素" + analysisElement);
                    System.out.println("tempSumofRow:" + tempSumofRow + "；num:" + num);
                    for (int row = tempSumofRow; row < num + tempSumofRow; row++) {
                        HashMap<String, Double> testElement = elementMassData.get(row - 1);
                        System.out.println("目前遞迴測試的元素數值:" + testElement.get(analysisElement));
                        if (testElement.containsKey(analysisElement)) {
                            accumulation += testElement.get(analysisElement);
                            System.out.println("累積含量" + accumulation);
                        }
                    }
                    Double elementWeightPercent = accumulation / currentSum;
                    System.out.println("重量百分比:" + elementWeightPercent);
                    Row rowOfElement = newSheet.createRow(tempWriteRow);
                    Cell cellOfElement = rowOfElement.createCell(1);
                    cellOfElement.setCellValue(analysisElement);
                    cellOfElement.setCellStyle(percentageStyle);
                    Cell cellOfWeight = rowOfElement.createCell(2);
                    cellOfWeight.setCellValue(elementWeightPercent);
                    cellOfWeight.setCellStyle(percentageStyle);
                    try (OutputStream fileOut = new FileOutputStream(urlOfExcel)) {
                        workbook.write(fileOut);
                        System.out.println("新增資料行數:" + tempWriteRow);
                    } catch (Exception e) {
                        e.printStackTrace();
                    }
                    tempWriteRow++;
                }
                writeRow += num;
                tempSumofRow += num;
            }
            workbook.close();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
