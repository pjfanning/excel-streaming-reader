package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.impl.adapter.RowAdapter;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.CellType;
import java.util.Map;
import java.util.TreeMap;
import java.util.Iterator;
import java.util.Spliterator;
import java.util.Spliterators;

public class StreamingRow implements RowAdapter {
  private final Sheet sheet;
  private final int rowIndex;
  private final boolean isHidden;
  private float height = -1.0f;
  private final TreeMap<Integer, Cell> cellMap;
  private StreamingSheetReader streamingSheetReader;
  private CellStyle rowStyle;

  public StreamingRow(Sheet sheet, int rowIndex, boolean isHidden) {
    this.sheet = sheet;
    this.rowIndex = rowIndex;
    this.isHidden = isHidden;
    cellMap = new TreeMap<>();
  }

  void setStreamingSheetReader(StreamingSheetReader streamingSheetReader) {
    this.streamingSheetReader = streamingSheetReader;
  }

  void setHeight(float height) {
    this.height = height;
  }

  public Map<Integer, Cell> getCellMap() {
    return cellMap;
  }

  /* Supported */

  /**
   * Get row number this row represents
   *
   * @return the row number (0 based)
   */
  @Override
  public int getRowNum() {
    return rowIndex;
  }

  /**
   * @return Cell iterator of the physically defined cells for this row.
   */
  @Override
  public Iterator<Cell> cellIterator() {
    return cellMap.values().iterator();
  }

  /**
   * @return Cell iterator of the physically defined cells for this row.
   */
  @Override
  public Iterator<Cell> iterator() {
    return cellMap.values().iterator();
  }

  @Override
  public Spliterator<Cell> spliterator() {
    return Spliterators.spliterator(cellMap.values(), Spliterator.ORDERED);
  }

  @Override
  public Sheet getSheet() {
    return sheet;
  }

  /**
   * Get the cell representing a given column (logical cell) 0-based.  If you
   * ask for a cell that is not defined, you get a null.
   *
   * @param cellnum 0 based column number
   * @return Cell representing that column or null if undefined.
   */
  @Override
  public Cell getCell(int cellnum) {
    return cellMap.get(cellnum);
  }

  /**
   * Gets the index of the last cell contained in this row <b>PLUS ONE</b>.
   *
   * @return short representing the last logical cell in the row <b>PLUS ONE</b>,
   * or -1 if the row does not contain any cells.
   */
  @Override
  public short getLastCellNum() {
    return (short) (cellMap.isEmpty() ? -1 : cellMap.lastEntry().getValue().getColumnIndex() + 1);
  }

  /**
   * Get whether or not to display this row with 0 height
   *
   * @return - zHeight height is zero or not.
   */
  @Override
  public boolean getZeroHeight() {
    return isHidden;
  }

  @Override
  public short getHeight() {
    return (short)(getHeightInPoints()*20);
  }

  @Override
  public float getHeightInPoints() {
    return height;
  }

  /**
   * Gets the number of defined cells (NOT number of cells in the actual row!).
   * That is to say if only columns 0,4,5 have values then there would be 3.
   *
   * @return int representing the number of defined cells in the row.
   */
  @Override
  public int getPhysicalNumberOfCells() {
    return cellMap.size();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public short getFirstCellNum() {
    if(cellMap.isEmpty()) {
      return -1;
    }
    return cellMap.firstKey().shortValue();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Cell getCell(int cellnum, MissingCellPolicy policy) {
    StreamingCell cell = (StreamingCell) cellMap.get(cellnum);
    if(policy == MissingCellPolicy.CREATE_NULL_AS_BLANK) {
      if(cell == null) {
        boolean use1904Dates = streamingSheetReader != null && streamingSheetReader.isUse1904Dates();
        return new StreamingCell(sheet, cellnum, this, use1904Dates);
      }
    } else if(policy == MissingCellPolicy.RETURN_BLANK_AS_NULL) {
      if(cell == null || cell.getCellType() == CellType.BLANK) { return null; }
    }
    return cell;
  }


  @Override
  public boolean isFormatted() {
    return rowStyle != null;
  }

  @Override
  public CellStyle getRowStyle() {
    return rowStyle;
  }

  @Override
  public void setRowStyle(CellStyle style) {
    this.rowStyle = style;
  }

}
