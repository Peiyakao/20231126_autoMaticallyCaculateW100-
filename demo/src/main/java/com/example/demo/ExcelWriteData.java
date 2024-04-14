package com.example.demo;

import java.lang.reflect.Array;

import org.hibernate.mapping.List;

public class ExcelWriteData {

    public int currentCellOfNumber;
    public String particleNameOfcell;
    public String cellOfElement;
    public Double cellOfWeight;

    public Integer getCurrentCellOfNumber() {
        return currentCellOfNumber;
    }

    public void setCurrentCellOfNumber(Integer currentCellOfNumber) {
        this.currentCellOfNumber = currentCellOfNumber;
    }

    public String getParticleNameOfcell() {
        return particleNameOfcell;
    }

    public void setParticleNameOfcell(String particleNameOfcell) {
        this.particleNameOfcell = particleNameOfcell;
    }

    public String getCellOfElement() {
        return cellOfElement;
    }

    public void setCellOfElement(String cellOfElement) {
        this.cellOfElement = cellOfElement;
    }

    public Double getCellOfWeight() {
        return cellOfWeight;
    }

    public void setCellOfWeight(Double cellOfWeight) {
        this.cellOfWeight = cellOfWeight;
    }
}
