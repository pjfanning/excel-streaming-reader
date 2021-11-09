package com.github.pjfanning.xlsx.impl;

import org.apache.poi.ss.formula.EvaluationCell;
import org.apache.poi.ss.formula.EvaluationSheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.util.Internal;

/**
 * wrapper for a sheet under evaluation
 */
@Internal
final class OoxmlEvaluationSheet implements EvaluationSheet {
  private final Sheet _xs;

  OoxmlEvaluationSheet(Sheet sheet) {
    _xs = sheet;
  }

  Sheet getSXSSFSheet() {
    return _xs;
  }

  /* (non-Javadoc)
   * @see org.apache.poi.ss.formula.EvaluationSheet#getlastRowNum()
   * @since POI 4.0.0
   */
  @Override
  public int getLastRowNum() {
    return _xs.getLastRowNum();
  }

  /* (non-Javadoc)
   * @see org.apache.poi.ss.formula.EvaluationSheet#isRowHidden(int)
   * @since POI 4.1.0
   */
  @Override
  public boolean isRowHidden(int rowIndex) {
    Row row = _xs.getRow(rowIndex);
    if (row == null) return false;
    return row.getZeroHeight();
  }

  @Override
  public EvaluationCell getCell(int rowIndex, int columnIndex) {
    Row row = _xs.getRow(rowIndex);
    if (row == null) {
      return null;
    }
    Cell cell = row.getCell(columnIndex);
    if (cell == null) {
      return null;
    }
    return new OoxmlEvaluationCell(cell, this);
  }

  /* (non-JavaDoc), inherit JavaDoc from EvaluationSheet
   * @since POI 3.15 beta 3
   */
  @Override
  public void clearAllCachedResultValues() {
  }
}
