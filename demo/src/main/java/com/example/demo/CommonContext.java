package com.example.demo;

public class CommonContext {
    public static Integer ANALYSIS_LOG_MESSAGE_ROW;
    public static String ANALYSIS_LOG_MESSAGE_CELL;
    public static String ANALYSIS_LOG_MESSAGE_ERROR_TYPE;
    public static String ANALYSIS_LOG_MESSAGE_EXCEPTION;
    public static String ANALYSIS_LOG_MESSAGE;
    public static String ANALYSIS_LOG_CURRENT_WORKSHEET;
    public static String ANALYSIS_LOG_START_MESSAGE;
    public static String ANALYSIS_LOG_END_MESSAGE;
    public static String ANALYSIS_LOG_EXCEL_UNCLOSED = "EXCEL開啟無法寫入重量百分比資訊";
    public static String ANALYSIS_TARGET_URL;
    public static Boolean INSERTTITLETOEXECL;
    public static Boolean INSERTRESULTTOEXCEL;

    public static void updateAnalysisLogMessage() {
        ANALYSIS_LOG_MESSAGE = "錯誤欄位:" + ANALYSIS_LOG_MESSAGE_CELL + ANALYSIS_LOG_MESSAGE_ROW
                + ANALYSIS_LOG_MESSAGE_ERROR_TYPE;
    }
}
