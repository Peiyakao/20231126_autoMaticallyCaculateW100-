package com.example.demo;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.HashMap;
import java.util.Objects;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.xssf.usermodel.*;
import java.nio.file.*;
import java.text.DecimalFormat;
import java.time.LocalDate;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.time.format.DateTimeParseException;
import java.util.Scanner;

public class GetExcelData {

    public ExcelParameter excelParameter = new ExcelParameter();

    public void getExcelUrl() {
        String urlOfExcel;
        Scanner inputUrl = new Scanner(System.in);
        System.out.println("請輸入分析檔案路徑:");
        urlOfExcel = inputUrl.nextLine();

        // 判斷輸入內容是否為路徑
        if (urlOfExcel.startsWith("\"") && urlOfExcel.endsWith("\"")) {
            urlOfExcel = urlOfExcel.substring(1, urlOfExcel.length() - 1);
        }
        Path pathofExcel = Paths.get(urlOfExcel);
        CommonContext.ANALYSIS_TARGET_URL = urlOfExcel;
        if (Files.exists(pathofExcel)) {
            excelParameter.setUrlOfExcel(urlOfExcel);
            // inputUrl.close();
            // Step 2 輸入工作表位置
            // 取得該路徑的log檔案情況
            Path parentDirectory = Paths.get(urlOfExcel).getParent();
            Path newFilePath = parentDirectory.resolve("analysisLog.xlsx");
            LocalDate today = LocalDate.now();
            String todayString = today.toString();
            if (Files.exists(newFilePath)) {
                try {
                    Workbook workbook = WorkbookFactory.create(Files.newInputStream(newFilePath));
                    int numberOfSheets = workbook.getNumberOfSheets();
                    String sheetName = workbook.getSheetName(numberOfSheets - 1);
                    if (sheetName.contains(todayString)) {
                        sheetName = sheetName.replace(todayString, "");
                        Integer newSheetNameIndex = Integer.parseInt(sheetName);
                        CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET = todayString + (newSheetNameIndex + 1);

                    } else {
                        CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET = todayString + 1;
                    }
                } catch (Exception e) {
                    CommonContext.ANALYSIS_LOG_MESSAGE_EXCEPTION = "取得分析檔案路徑log檔案異常" + e;
                    analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE_EXCEPTION);
                }

            } else {
                CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET = todayString + 0;
            }

            getExcelWorkSheet();

        } else {
            System.out.println("查無所輸入檔案，請重新輸入");
            getExcelUrl();
        }

    }

    public void getExcelWorkSheet() {
        Integer workSheet;
        String exit;
        Scanner inputworkSheet = new Scanner(System.in);
        System.out.println("請輸入欲分析工作表的序位，例如:第一張工作表請填1:");
        if (inputworkSheet.hasNextLine()) {
            exit = inputworkSheet.nextLine();
            if (exit.equals("back")) {
                getExcelUrl();
            } else {
                try {
                    workSheet = Integer.parseInt(exit);
                    if (workSheet >= 1) {
                        excelParameter.setWorkSheet(workSheet - 1);
                        getExcelColumnOfSum();
                    } else {
                        System.out.println("請輸入正整數！");
                        getExcelWorkSheet();
                    }
                } catch (NumberFormatException e) {
                    System.out.println("請輸入正整數！");
                    getExcelWorkSheet();
                }
            }

        }
        // inputworkSheet.close();

    }

    public void getExcelColumnOfSum() {
        String columnOfSumString;
        Integer columnOfSum = 0;
        Scanner inputSum = new Scanner(System.in);
        System.out.println("請輸入部件/成品(Article)重量欄位，如果是A欄填寫A:");
        if (inputSum.hasNextLine()) {
            columnOfSumString = inputSum.nextLine();
            if (columnOfSumString.equals("back")) {
                getExcelWorkSheet();
            }
            if (columnOfSumString.matches("[A-Za-z]+")) {
                for (int i = 0; i < columnOfSumString.length(); i++) {
                    columnOfSum = columnOfSum + (int) columnOfSumString.toUpperCase().charAt(i) - 64;
                }
                excelParameter.setColumnOfSum(columnOfSum - 1);
                getExcelColumnOfElement();
            } else {
                System.out.println("請輸入部件/成品(Article)重量欄位，如果是A欄填寫A:");
                getExcelColumnOfSum();
            }
        } else {
            System.out.println("請輸入部件/成品(Article)重量欄位，如果是A欄填寫A:");
            getExcelColumnOfSum();
        }
        // inputSum.close();

    }

    public void getExcelColumnOfElement() {
        String columnOfElementString;
        Integer columnOfElement = 0;
        Scanner inputElement = new Scanner(System.in);
        System.out.println("請輸入CAS Number欄位，如果是A欄填寫A:");

        // inputElement.close();
        if (inputElement.hasNextLine()) {
            columnOfElementString = inputElement.nextLine();
            if (columnOfElementString.equals("back")) {
                getExcelColumnOfSum();
            }
            if (columnOfElementString.matches("[A-Za-z]+")) {
                for (int i = 0; i < columnOfElementString.length(); i++) {
                    columnOfElement = columnOfElement + (int) columnOfElementString.toUpperCase().charAt(i) - 64;
                }
                excelParameter.setColumnOfElement(columnOfElement - 1);
                getExcelColumnOfMass();
            } else {
                System.out.println("請輸入CAS Number欄位，如果是A欄填寫A:");
                getExcelColumnOfElement();
            }
        } else {
            System.out.println("請輸入CAS Number欄位，如果是A欄填寫A:");
            getExcelColumnOfElement();
        }

    }

    public void getExcelColumnOfMass() {
        String columnOfMassString;
        Integer columnOfMass = 0;
        Scanner inputMass = new Scanner(System.in);
        System.out.println("請輸入物質重量欄位，如果是A欄填寫A:");
        // inputMass.close();
        if (inputMass.hasNextLine()) {
            columnOfMassString = inputMass.nextLine();
            if (columnOfMassString.equals("back")) {
                getExcelColumnOfElement();
            }
            if (columnOfMassString.matches("[A-Za-z]+")) {
                for (int i = 0; i < columnOfMassString.length(); i++) {

                    columnOfMass = columnOfMass + (int) columnOfMassString.toUpperCase().charAt(i) - 64;
                }
                excelParameter.setColumnOfMass(columnOfMass - 1);
                getExcelColumnOfRange();
            } else {
                System.out.println("請輸入物質重量欄位，如果是A欄填寫A:");
                getExcelColumnOfMass();
            }
        } else {
            System.out.println("請輸入物質重量欄位，如果是A欄填寫A:");
            getExcelColumnOfMass();
        }

    }

    public void getExcelAnalysisSheet() {
        String workSheet;
        Scanner inputAnalysisSheet = new Scanner(System.in);

        HashMap<String, String> nameOfElement = new HashMap<>();
        DecimalFormat decimalFormat = new DecimalFormat("#");
        System.out.println("請輸入欲分析物質/禁限用物質列表檔案路徑:");
        workSheet = inputAnalysisSheet.nextLine();
        // 判斷輸入內容是否為路徑
        if (workSheet.startsWith("\"") && workSheet.endsWith("\"")) {
            workSheet = workSheet.substring(1, workSheet.length() - 1);
        }
        Path pathofExcel = Paths.get(workSheet);
        if (Files.exists(pathofExcel)) {
            excelParameter.setElementOfAnaylis(workSheet);
            Workbook workbook = null;
            try {
                // InputStream excelFile = new FileInputStream(workSheet);
                if (workSheet.contains("xlsx")) {
                    workbook = new XSSFWorkbook(new FileInputStream(workSheet));

                } else {
                    workbook = new HSSFWorkbook(new FileInputStream(workSheet));
                }
                Sheet sheet = workbook.getSheetAt(0);
                // Workbook workbook = WorkbookFactory.create(excelFile);
                // CommonContext.ANALYSIS_LOG_START_MESSAGE = "開始讀取分析物質/禁限用物質列表";
                // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
                // CommonContext.ANALYSIS_LOG_START_MESSAGE);
                for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                    Row sumRow = sheet.getRow(rowNum);
                    Cell cellOfId = sumRow.getCell(0);
                    Cell cellOfName = sumRow.getCell(1);
                    String cellOfNameString = "";
                    String cellOfIdString = "";
                    if (cellOfName.getCellType() == CellType.NUMERIC
                            && cellOfId.getCellType() == CellType.NUMERIC) {
                        cellOfNameString = decimalFormat.format(cellOfName.getNumericCellValue());
                        cellOfIdString = decimalFormat.format(cellOfId.getNumericCellValue());
                    } else if (cellOfName.getCellType() == CellType.NUMERIC
                            || cellOfId.getCellType() == CellType.NUMERIC) {
                        if (cellOfName.getCellType() == CellType.NUMERIC) {
                            cellOfNameString = decimalFormat.format(cellOfName.getNumericCellValue());
                            cellOfIdString = cellOfId.getStringCellValue();
                        } else if (cellOfId.getCellType() == CellType.NUMERIC) {
                            cellOfIdString = decimalFormat.format(cellOfId.getNumericCellValue());
                            cellOfNameString = cellOfName.getStringCellValue();
                        }
                    } else if (cellOfName.getCellType() == CellType.STRING) {
                        // if (cellOfId.getStringCellValue() == "") {
                        // CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                        // CommonContext.ANALYSIS_LOG_MESSAGE_CELL = "A";
                        // CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "空值";
                        // CommonContext.updateAnalysisLogMessage();
                        // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
                        // CommonContext.ANALYSIS_LOG_MESSAGE);
                        // }
                        // if (cellOfName.getStringCellValue() == "") {
                        // CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                        // CommonContext.ANALYSIS_LOG_MESSAGE_CELL = "B";
                        // CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "空值";
                        // CommonContext.updateAnalysisLogMessage();
                        // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
                        // CommonContext.ANALYSIS_LOG_MESSAGE);

                        // }

                        cellOfIdString = cellOfId.getStringCellValue();
                        cellOfNameString = cellOfName.getStringCellValue();
                    }
                    nameOfElement.put(cellOfIdString, cellOfNameString);
                    System.out.println(cellOfIdString + "," + cellOfNameString);

                }
                // CommonContext.ANALYSIS_LOG_END_MESSAGE = "讀取分析物質/禁限用物質列表結束";
                // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
                // CommonContext.ANALYSIS_LOG_END_MESSAGE);
                excelParameter.setNameOfElement(nameOfElement);
                getUrlOfElementName();

            } catch (Exception e) {
                CommonContext.ANALYSIS_LOG_MESSAGE_EXCEPTION = "讀取分析物質/禁限用物質列表異常" + e;
                analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE_EXCEPTION);
            }

            // inputUrl.close();
            // Step 2 輸入工作表位置

        } else {
            System.out.println("查無所輸入檔案，請重新輸入");
            getExcelAnalysisSheet();
        }

    }

    public void getExcelColumnOfRange() {
        String columnOfRangeExit;
        Integer columnOfRange = 1;
        Scanner inputRange = new Scanner(System.in);
        System.out.println("請輸入資料範圍，如果數值最後一列為第16列，請填16:");
        excelParameter.setColumnOfRange(columnOfRange);
        if (inputRange.hasNextLine()) {
            columnOfRangeExit = inputRange.nextLine();
            if (columnOfRangeExit.equals("back")) {
                getExcelColumnOfMass();
            }
            if (columnOfRangeExit.matches("[0-9]+")) {
                columnOfRange = Integer.parseInt(columnOfRangeExit);
                excelParameter.setColumnOfRange(columnOfRange);
                getExcelAnalysisSheet();
            } else {
                System.out.println("請輸入資料範圍，如果數值最後一列為第16列，請填16:");
                getExcelColumnOfRange();
            }
        } else {
            System.out.println("請輸入資料範圍，如果數值最後一列為第16列，請填16:");
            getExcelColumnOfRange();
        }

    }

    public void getUrlOfElementName() {
        Double weightPercentLimit;
        Scanner inputweightPercentLimit = new Scanner(System.in);
        System.out.println("請輸入物質濃度上限，如需將分析物質濃度全部列出，請輸入0:");
        String input = inputweightPercentLimit.nextLine();
        try {
            weightPercentLimit = Double.parseDouble(input);
            if (weightPercentLimit <= 100 && weightPercentLimit >= 0) {
                excelParameter.setWeightPercentLimit(weightPercentLimit);
                excelParameter.setElementOfAnaylis(excelParameter.getElementOfAnaylis());
                getElementOfAnaylis(excelParameter.getElementOfAnaylis());

            } else {
                System.out.println("請輸入0-100數字");
                getExcelAnalysisSheet();
            }
        } catch (Exception e) {
            System.out.println("請輸入0-100數字");
            getExcelAnalysisSheet();
        }

    }

    public void getElementOfAnaylis(String analysisUrl) throws Exception {
        ArrayList<String> analysisList = new ArrayList();
        DecimalFormat decimalFormat = new DecimalFormat("#");
        Workbook workbook = null;
        try {
            // InputStream excelFile = new FileInputStream(analysisUrl);
            if (analysisUrl.contains("xlsx")) {
                workbook = new XSSFWorkbook(new FileInputStream(analysisUrl));
            } else {
                workbook = new HSSFWorkbook(new FileInputStream(analysisUrl));
            }
            // Workbook workbook = WorkbookFactory.create(excelFile);
            // CommonContext.ANALYSIS_LOG_START_MESSAGE = "開始讀取分析物質";
            // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
            // CommonContext.ANALYSIS_LOG_START_MESSAGE);
            Sheet sheet = workbook.getSheetAt(0);
            for (int rowNum = 1; rowNum <= sheet.getLastRowNum(); rowNum++) {
                Row sumRow = sheet.getRow(rowNum);
                Cell cell = sumRow.getCell(0);
                String cellString = "";
                if (cell.getCellType() == CellType.NUMERIC) {
                    cellString = decimalFormat.format(cell.getNumericCellValue());
                } else if (cell.getCellType() == CellType.STRING) {

                    cellString = String.valueOf(cell.getStringCellValue());

                }
                analysisList.add(cellString.trim());
            }
            excelParameter.setElementOfAnaylist(analysisList);
            System.out.println("本次分析元素物質為:" + String.join(",", analysisList));
            // CommonContext.ANALYSIS_LOG_END_MESSAGE = "讀取分析物質結束";
            // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
            // CommonContext.ANALYSIS_LOG_END_MESSAGE);
            getExcelDataSumOfNum(excelParameter.getUrlOfExcel(),
                    excelParameter.getWorkSheet(),
                    excelParameter.getColumnOfSum(),
                    excelParameter.getColumnOfRange());
        } catch (Exception e) {
            CommonContext.ANALYSIS_LOG_MESSAGE_EXCEPTION = "讀取分析物質異常" + e;
            analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE_EXCEPTION);
        }

    }

    // 創建一個分析資料的log紀錄excel檔案
    public void analysisLogExcel(String pathOfAnalysis, String logMessage) {
        // 取得父層路徑
        Path parentDirectory = Paths.get(pathOfAnalysis).getParent();
        Path newFilePath = parentDirectory.resolve("analysisLog.xlsx");
        if (Files.exists(newFilePath)) {
            try (Workbook workbook = WorkbookFactory.create(Files.newInputStream(newFilePath))) {
                Sheet sheet;
                int sheetIndex = workbook.getSheetIndex(CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET);
                if (sheetIndex == -1) {
                    sheet = workbook.createSheet(CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET);
                    Row rowOfTitle = sheet.createRow(0);
                    Cell cellOfNo = rowOfTitle.createCell(0);
                    cellOfNo.setCellValue("No");
                    Cell cellOfMessage = rowOfTitle.createCell(1);
                    cellOfMessage.setCellValue("Message");
                } else {
                    sheet = workbook.getSheet(CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET);
                }
                int lastRowIndex = sheet.getLastRowNum();
                Row newRow = sheet.createRow(lastRowIndex + 1);
                Cell cellNo = newRow.createCell(0);
                cellNo.setCellValue(lastRowIndex + 1);
                Cell cellOfMessage = newRow.createCell(1);
                cellOfMessage.setCellValue(logMessage);
                try {
                    workbook.write(Files.newOutputStream(newFilePath));
                    System.out.println("寫入一筆log紀錄：");
                } catch (java.io.IOException e) {
                    System.out.println("無法寫入一筆log紀錄：" + e.getMessage());
                }
            } catch (Exception e) {
                System.out.println("無法開啟既有log紀錄表：" + e.getMessage());
            }

        } else {
            try (Workbook workbook = new XSSFWorkbook()) {
                System.out.println(CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET);
                Sheet sheet = workbook.createSheet(CommonContext.ANALYSIS_LOG_CURRENT_WORKSHEET);
                Row rowOfTitle = sheet.createRow(0);
                Cell cellOfNo = rowOfTitle.createCell(0);
                cellOfNo.setCellValue("No");
                Cell cellOfMessage = rowOfTitle.createCell(1);
                cellOfMessage.setCellValue("Message");
                Row rowOfMessage = sheet.createRow(1);
                Cell cellOfMessageNo = rowOfMessage.createCell(0);
                cellOfMessageNo.setCellValue(1);
                Cell cellOfMessageContent = rowOfMessage.createCell(1);
                cellOfMessageContent.setCellValue(logMessage);

                try {
                    workbook.write(Files.newOutputStream(newFilePath));
                    System.out.println("Excel 文件已成功建立新Log資訊表");
                } catch (java.io.IOException e) {
                    System.out.println("寫入文件失敗：" + e.getMessage());
                }
            } catch (java.io.IOException e) {
                System.out.println("新增Log文件失敗：" + e.getMessage());
            }
        }
    }

    public void getExcelDataSumOfNum(String urlOfExcel, Integer workSheet, Integer columnOfSum,
            Integer columnOfRange) {
        // 載入EXCEL資料
        ArrayList<Integer> numOfParticle = new ArrayList();
        Workbook workbook = null;
        try {
            if (urlOfExcel.contains("xlsx")) {
                workbook = new XSSFWorkbook(new FileInputStream(urlOfExcel));
            } else {
                workbook = new HSSFWorkbook(new FileInputStream(urlOfExcel));
            }
            // Workbook workbook = WorkbookFactory.create(excelFile);
            Sheet sheet = workbook.getSheetAt(workSheet);
            // System.out.println("讀取最小物質數量");
            Integer temp = 1;
            for (int rowNum = 2; rowNum < columnOfRange; rowNum++) {
                Row sumRow = sheet.getRow(rowNum);
                Cell cell = sumRow.getCell(columnOfSum);
                CellType cellType = cell.getCellType();
                if (cellType == CellType.BLANK) {
                    temp++;

                } else if (cellType == CellType.NUMERIC) {
                    numOfParticle.add(temp);
                    temp = 1;
                }
                if (rowNum == columnOfRange - 1) {
                    numOfParticle.add(temp);
                    temp = 1;
                    workbook.close();
                }
            }

        } catch (Exception e) {
            e.printStackTrace();
        }
        excelParameter.setNumOfParticle(numOfParticle);
        getElementMassData(
                excelParameter.getUrlOfExcel(),
                excelParameter.getWorkSheet(),
                excelParameter.getColumnOfElement(),
                excelParameter.getColumnOfMass(),
                excelParameter.getColumnOfRange());
    }

    public ArrayList<String> getAnalysisEleList(String element) {
        String[] strArray = element.split(",");
        ArrayList<String> analysisElement = new ArrayList<>(Arrays.asList(strArray));
        return analysisElement;

    }

    public void getElementMassData(String urlOfExcel, Integer workSheet,
            Integer columnOfElement,
            Integer columnOfMass,
            Integer columnOfRange) {
        ArrayList<HashMap<String, Double>> elementMassData = new ArrayList<>();
        Integer isAnyError = 0;
        Workbook workbook = null;
        try {
            if (urlOfExcel.contains("xlsx")) {
                workbook = new XSSFWorkbook(new FileInputStream(urlOfExcel));
            } else {
                workbook = new HSSFWorkbook(new FileInputStream(urlOfExcel));
            }
            // Workbook workbook = WorkbookFactory.create(excelFile);
            Sheet sheet = workbook.getSheetAt(workSheet);
            System.out.println("讀取物質名稱及數值");
            // CommonContext.ANALYSIS_LOG_START_MESSAGE = "開始讀取物質名稱及數值";
            // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
            // CommonContext.ANALYSIS_LOG_START_MESSAGE);
            for (int rowNum = 1; rowNum < columnOfRange; rowNum++) {
                Row elementAndMassRow = sheet.getRow(rowNum);
                Cell element = elementAndMassRow.getCell(columnOfElement);
                Cell mass = elementAndMassRow.getCell(columnOfMass);
                HashMap<String, Double> tempElement = new HashMap<>();
                // 先判定欄位名稱符合規則在判斷重量欄位
                if (element.getCellType() == CellType.NUMERIC) {
                    String cas = String.valueOf(element.getNumericCellValue()).trim();
                    System.out.println("NUM" + cas);
                    // 欄位名稱日期格式不做事
                    try {
                        Date isDate = element.getDateCellValue();
                        CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                        CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfElement);
                        CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "欄位物質名稱為日期格式";
                        CommonContext.updateAnalysisLogMessage();
                        analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);

                    } catch (Exception e) {

                    }
                    if (mass.getCellType() == CellType.NUMERIC || mass.getCellType() == CellType.FORMULA) {
                        try {
                            tempElement.put(cas, mass.getNumericCellValue());
                        } catch (Exception e) {
                            CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                            CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfMass);
                            CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "欄位重量數值型態異常，請檢查是否為公式";
                            CommonContext.updateAnalysisLogMessage();
                            analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                            tempElement.put(cas, -1.00);
                        }
                    } else if (mass.getCellType() == CellType.STRING) {
                        CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                        CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfMass);
                        CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "重量數值型態為文字或空白";
                        CommonContext.updateAnalysisLogMessage();
                        analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                        tempElement.put(cas, -1.00);
                    } else {
                        System.out.println("第" + (rowNum + 1) + "行欄位重量數值型態有誤");

                        // 將欄位index轉回英文
                        CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                        CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfMass);
                        CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "重量數值型態非數值";
                        CommonContext.updateAnalysisLogMessage();
                        analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                        tempElement.put(cas, -1.00);
                    }

                } else if (element.getCellType() == CellType.STRING) {
                    System.out.println("STR" + element.getStringCellValue());

                    if (mass.getCellType() == CellType.NUMERIC || mass.getCellType() == CellType.FORMULA) {
                        try {
                            tempElement.put(element.getStringCellValue().trim(), mass.getNumericCellValue());
                        } catch (Exception e) {
                            CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                            CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfMass);
                            CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "欄位重量數值型態異常，請檢查是否為公式";
                            CommonContext.updateAnalysisLogMessage();
                            analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                            tempElement.put(element.getStringCellValue().trim(), -1.00);
                        }

                    } else if (mass.getCellType() == CellType.STRING) {
                        CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                        CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfMass);
                        CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "重量數值型態為文字或空白";
                        CommonContext.updateAnalysisLogMessage();
                        analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                        tempElement.put(element.getStringCellValue().trim(), -1.00);
                    } else {
                        System.out.println("第" + (rowNum + 1) + "行欄位重量數值型態有誤");
                        // 將欄位index轉回英文
                        CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                        CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfMass);
                        CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "重量數值型態非數值";
                        CommonContext.updateAnalysisLogMessage();
                        analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                        tempElement.put(element.getStringCellValue().trim(), -1.00);
                    }

                    elementMassData.add(tempElement);
                } else if (element.getCellType() == CellType.BLANK) {
                    CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                    CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfElement);
                    CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "欄位物質名稱為空值";
                    CommonContext.updateAnalysisLogMessage();
                    analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                } else if (element.getCellType() == CellType.FORMULA) {
                    CommonContext.ANALYSIS_LOG_MESSAGE_ROW = rowNum + 1;
                    CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfElement);
                    CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "欄位物質名稱為公式";
                    CommonContext.updateAnalysisLogMessage();
                    analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, CommonContext.ANALYSIS_LOG_MESSAGE);
                }
            }
        } catch (Exception e) {
            analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL, "取得分析表物質質量意外錯誤");
        }

        // CommonContext.ANALYSIS_LOG_END_MESSAGE = "讀取物質名稱及數值結束";
        // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
        // CommonContext.ANALYSIS_LOG_END_MESSAGE);

        if (isAnyError == 0) {
            excelParameter.setElementMassData(elementMassData);
            calculateWeightPercent(
                    excelParameter.getNumOfParticle(),
                    excelParameter.getElementOfAnaylist(),
                    excelParameter.getElementMassData(),
                    excelParameter.getNameOfElement(),
                    excelParameter.getUrlOfExcel(),
                    excelParameter.getWorkSheet(),
                    excelParameter.getColumnOfSum());
        } else {
            System.out.println("欄位重量數值型態有誤，停止分析，執行ctrl+c離開程式");
        }

    }

    public String exchangeIndexToChar(Integer index) {
        String indexToChar = "";
        char tempChar = (char) (index + 65);
        indexToChar = String.valueOf(tempChar);
        return indexToChar;

    };

    public void calculateWeightPercent(
            ArrayList<Integer> numOfParticle,
            ArrayList<String> elementOfAnaylis,
            ArrayList<HashMap<String, Double>> elementMassData,
            HashMap<String, String> nameOfElement,
            String urlOfExcel,
            Integer workSheet,
            Integer columnOfSum) {
        Integer tempSumofRow = 1;
        Integer writeRow = 1;
        Workbook workbook = null;
        try {
            // Workbook workbook = WorkbookFactory.create(excelFile);
            if (urlOfExcel.contains("xlsx")) {
                workbook = new XSSFWorkbook(new FileInputStream(urlOfExcel));
            } else {
                workbook = new HSSFWorkbook(new FileInputStream(urlOfExcel));
            }
            Sheet sheet = workbook.getSheetAt(workSheet);
            Sheet newSheet = workbook.createSheet();
            Row rowOfTitle = newSheet.createRow(0);
            Cell cellOfNum = rowOfTitle.createCell(0);
            cellOfNum.setCellValue("編號");
            Cell cellOfIngNum = rowOfTitle.createCell(1);
            cellOfIngNum.setCellValue("部件料號");
            Cell cellOfPaticName = rowOfTitle.createCell(2);
            cellOfPaticName.setCellValue("物質名稱");
            Cell cellOfPW = rowOfTitle.createCell(3);
            cellOfPW.setCellValue("物質含量百分比");
            CellStyle percentageStyle = workbook.createCellStyle();
            percentageStyle.setDataFormat(workbook.createDataFormat().getFormat("0.00%"));
            try (OutputStream fileOut = new FileOutputStream(urlOfExcel)) {
                workbook.write(fileOut);
                System.out.println("新增表頭");

            } catch (Exception e) {
                e.printStackTrace();
            }
            Integer numberOfElement = 0;
            Integer currentCellOfNumber = 0;
            String currentCellOfPartName = "";
            Integer totalAddRow = 0;
            // CommonContext.ANALYSIS_LOG_START_MESSAGE = "開始計算物質重量百分比";
            // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
            // CommonContext.ANALYSIS_LOG_START_MESSAGE);
            for (Integer num : numOfParticle) {
                String accumulationError = "";
                Row sumRow = sheet.getRow(tempSumofRow);
                Cell cell = sumRow.getCell(columnOfSum);
                Double currentSum = cell.getNumericCellValue();
                Integer tempWriteRow = writeRow;
                for (String analysisElement : elementOfAnaylis) {
                    Double accumulation = 0.00;
                    // System.out.println("目前分析元素" + analysisElement);
                    System.out.println(elementMassData.size());
                    for (int row = tempSumofRow; row < num + tempSumofRow; row++) {
                        HashMap<String, Double> testElement = elementMassData.get(row - 1);
                        // System.out.println(analysisElement);
                        // System.out.println(testElement);
                        if (testElement.containsKey(analysisElement)) {
                            try {
                                System.out.println(testElement.get(analysisElement));
                                if (testElement.get(analysisElement) != -1) {
                                    accumulation += testElement.get(analysisElement);
                                    System.out.println(analysisElement + "累積含量" + accumulation);
                                } else {
                                    accumulationError = "ERROR";
                                    System.out.println("ERROR");
                                }
                            } catch (Exception e) {
                                System.out.println("第" + row + "行欄位重量數值型態有誤");
                                CommonContext.ANALYSIS_LOG_MESSAGE_ROW = row;
                                CommonContext.ANALYSIS_LOG_MESSAGE_CELL = exchangeIndexToChar(columnOfSum);
                                CommonContext.ANALYSIS_LOG_MESSAGE_ERROR_TYPE = "欄位總重量數型態不為數值，請檢查是否為公式";
                                CommonContext.updateAnalysisLogMessage();
                                analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
                                        CommonContext.ANALYSIS_LOG_MESSAGE);
                            }

                        }
                    }

                    Double elementWeightPercent = accumulation / currentSum;
                    System.out.println(elementWeightPercent + "=" + accumulation + "/" + currentSum);
                    Row rowOfElement = newSheet.createRow(tempWriteRow);
                    Cell cellOfNumber = rowOfElement.createCell(0);
                    Cell cellOfPartName = rowOfElement.createCell(1);
                    Row particleNumberOfRow = sheet.getRow(numberOfElement + 1);
                    Cell particleNumberOfcell = particleNumberOfRow.getCell(0);
                    Cell particleNameOfcell = particleNumberOfRow.getCell(1);
                    if (Objects.equals(tempWriteRow, numberOfElement + 1)) {
                        if (particleNumberOfcell.getNumericCellValue() > 0) {
                            currentCellOfNumber = (int) particleNumberOfcell.getNumericCellValue();
                            currentCellOfPartName = particleNameOfcell.getStringCellValue();
                            cellOfNumber.setCellValue(currentCellOfNumber);
                            cellOfPartName.setCellValue(currentCellOfPartName);
                            try (OutputStream fileOut = new FileOutputStream(urlOfExcel)) {
                                workbook.write(fileOut);
                            } catch (Exception e) {
                                e.printStackTrace();
                            }

                        } else if ((particleNumberOfRow.getCell(0) == null)
                                || (particleNumberOfcell.getNumericCellValue() == 0)) {
                            cellOfNumber.setCellValue(currentCellOfNumber);
                            cellOfPartName.setCellValue(currentCellOfPartName);
                            try (OutputStream fileOut = new FileOutputStream(urlOfExcel)) {
                                workbook.write(fileOut);
                            } catch (Exception e) {
                                e.printStackTrace();
                            }
                        }

                    }
                    if (accumulationError.equals("ERROR")) {
                        totalAddRow++;
                        Cell cellOfElement = rowOfElement.createCell(2);
                        cellOfElement.setCellValue(nameOfElement.get(analysisElement));
                        Cell cellOfWeight = rowOfElement.createCell(3);
                        cellOfWeight.setCellValue(accumulationError);
                        try (OutputStream fileOut = new FileOutputStream(urlOfExcel)) {
                            workbook.write(fileOut);
                            System.out.println("新增資料行數:" + tempWriteRow);
                        } catch (Exception e) {
                            analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
                                    CommonContext.ANALYSIS_LOG_EXCEL_UNCLOSED + e);
                        }
                    }
                    if (elementWeightPercent >= excelParameter.getWeightPercentLimit()
                            && accumulationError.equals("")) {
                        totalAddRow++;
                        Cell cellOfElement = rowOfElement.createCell(2);
                        cellOfElement.setCellValue(nameOfElement.get(analysisElement));
                        Cell cellOfWeight = rowOfElement.createCell(3);
                        cellOfWeight.setCellValue(elementWeightPercent);
                        cellOfWeight.setCellStyle(percentageStyle);
                        try (OutputStream fileOut = new FileOutputStream(urlOfExcel)) {
                            workbook.write(fileOut);
                            System.out.println("新增資料行數:" + tempWriteRow);
                        } catch (Exception e) {
                            analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
                                    CommonContext.ANALYSIS_LOG_EXCEL_UNCLOSED + e);
                        }
                    }
                    tempWriteRow++;
                }
                writeRow += num;
                tempSumofRow += num;
                numberOfElement += num;

            }

            if (totalAddRow > 0) {
                System.out.println("本次執行發現至少一筆物質含量超過規範");

            } else {
                System.out.println("本次執行無發現含量超過規範之物質");

            }
            // CommonContext.ANALYSIS_LOG_END_MESSAGE = "計算物質重量百分比結束";
            // analysisLogExcel(CommonContext.ANALYSIS_TARGET_URL,
            // CommonContext.ANALYSIS_LOG_END_MESSAGE);
            workbook.close();

        } catch (Exception e) {
            e.printStackTrace();
        }
    }

}
