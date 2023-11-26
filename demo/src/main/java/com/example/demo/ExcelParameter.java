package com.example.demo;

import java.util.ArrayList;
import java.util.HashMap;

public class ExcelParameter {

    String UrlOfExcel;
    ArrayList<Integer> numOfParticle;
    String elementOfAnaylis;

    public String getElementOfAnaylis() {
        return elementOfAnaylis;
    }

    public void setElementOfAnaylis(String elementOfAnaylis) {
        this.elementOfAnaylis = elementOfAnaylis;
    }

    public ArrayList<Integer> getNumOfParticle() {
        return numOfParticle;
    }

    public void setNumOfParticle(ArrayList<Integer> numOfParticle) {
        this.numOfParticle = numOfParticle;
    }

    public String getUrlOfExcel() {
        return UrlOfExcel;
    }

    public void setUrlOfExcel(String urlOfExcel) {
        UrlOfExcel = urlOfExcel;
    }

    public Integer getWorkSheet() {
        return WorkSheet;
    }

    public void setWorkSheet(Integer workSheet) {
        WorkSheet = workSheet;
    }

    public Integer getColumnOfSum() {
        return ColumnOfSum;
    }

    public void setColumnOfSum(Integer columnOfSum) {
        ColumnOfSum = columnOfSum;
    }

    public Integer getColumnOfMass() {
        return ColumnOfMass;
    }

    public void setColumnOfMass(Integer columnOfMass) {
        ColumnOfMass = columnOfMass;
    }

    Integer WorkSheet;
    Integer ColumnOfSum;
    Integer ColumnOfMass;
    Integer ColumnOfElement;
    Integer ColumnOfRange;

    public Integer getColumnOfRange() {
        return ColumnOfRange;
    }

    public void setColumnOfRange(Integer columnOfRange) {
        ColumnOfRange = columnOfRange;
    }

    public Integer getColumnOfElement() {
        return ColumnOfElement;
    }

    public void setColumnOfElement(Integer columnOfElement) {
        ColumnOfElement = columnOfElement;
    }

}
