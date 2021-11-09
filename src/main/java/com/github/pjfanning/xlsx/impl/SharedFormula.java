package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.util.CellAddress;

class SharedFormula {

  CellAddress cellAddress;
  String formula;

  SharedFormula(CellAddress cellAddress, String formula) {
    this.cellAddress = cellAddress;
    this.formula = formula;
  }

  CellAddress getCellAddress() {
    return cellAddress;
  }

  String getFormula() {
    return formula;
  }
}
