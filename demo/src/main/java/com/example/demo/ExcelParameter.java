package com.example.demo;

import java.util.ArrayList;
import java.util.HashMap;

import org.springframework.stereotype.Component;

@Component
public class ExcelParameter {
    Double weightPercentLimit;

    public Double getWeightPercentLimit() {
        return weightPercentLimit;
    }

    public void setWeightPercentLimit(Double weightPercentLimit) {
        this.weightPercentLimit = weightPercentLimit;
    }

    ArrayList<HashMap<String, Double>> elementMassData;

    public ArrayList<HashMap<String, Double>> getElementMassData() {
        return elementMassData;
    }

    public void setElementMassData(ArrayList<HashMap<String, Double>> elementMassData) {
        this.elementMassData = elementMassData;
    }

    String UrlOfExcel;
    ArrayList<Integer> numOfParticle;
    String elementOfAnaylis;
    HashMap<String, String> nameOfElement;
    ArrayList<String> elementOfAnaylist;

    public ArrayList<String> getElementOfAnaylist() {
        return elementOfAnaylist;
    }

    public void setElementOfAnaylist(ArrayList<String> elementOfAnaylist) {
        this.elementOfAnaylist = elementOfAnaylist;
    }

    public HashMap<String, String> getNameOfElement() {
        return nameOfElement;
    }

    public void setNameOfElement(HashMap<String, String> nameOfElement) {
        this.nameOfElement = nameOfElement;
    }

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
