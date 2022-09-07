package com.github.pjfanning.xlsx.impl;

import com.github.pjfanning.xlsx.CloseableIterator;
import com.github.pjfanning.xlsx.SharedFormula;
import com.github.pjfanning.xlsx.impl.adapter.SheetAdapter;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.PaneInformation;
import org.apache.poi.util.Units;
import org.apache.poi.xssf.model.Comments;
import org.apache.poi.xssf.usermodel.XSSFComment;

import java.util.*;

public class StreamingSheet implements SheetAdapter {

  private final String name;
  private final StreamingSheetReader reader;

  public StreamingSheet(String name, StreamingSheetReader reader) {
    this.name = name;
    this.reader = reader;
    reader.setSheet(this);
  }

  StreamingSheetReader getReader() {
    return reader;
  }

  /* Supported */

  /**
   * Workbook is only set under certain usage flows.
   */
  @Override
  public Workbook getWorkbook() {
    return reader.getWorkbook();
  }

  /**
   * Alias for {@link #rowIterator()} to allow foreach loops
   *
   * @return the streaming iterator, an instance of {@link CloseableIterator} -
   * it is recommended that you close the iterator when finished with it if you intend to keep the sheet open.
   */
  @Override
  public CloseableIterator<Row> iterator() {
    return reader.iterator();
  }

  /**
   * Returns a new iterator of the physical rows. This is an iterator of the PHYSICAL rows.
   * Meaning the 3rd element may not be the third row if say for instance the second row is undefined.
   *
   * This behaviour changed in v4.0.0. Earlier versions only created one simple iterator and repeated
   * calls to this method just returned the same iterator. Creating multiple iterators will slow down
   * your application and should be avoided unless necessary.
   *
   * @return the streaming iterator, an instance of {@link CloseableIterator} -
   * it is recommended that you close the iterator when finished with it if you intend to keep the sheet open.
   */
  @Override
  public CloseableIterator<Row> rowIterator() {
    return reader.iterator();
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public Spliterator<Row> spliterator() {
    // Long.MAX_VALUE is the documented value to use if the size is unknown
    return Spliterators.spliterator(rowIterator(), Long.MAX_VALUE, Spliterator.ORDERED);
  }

  /**
   * {@inheritDoc}
   */
  @Override
  public String getSheetName() {
    return name;
  }

  /**
   * Get the hidden state for a given column
   *
   * @param columnIndex - the column to set (0-based)
   * @return hidden - <code>false</code> if the column is visible
   */
  @Override
  public boolean isColumnHidden(int columnIndex) {
    return reader.isColumnHidden(columnIndex);
  }

  /**
   * Gets the first row on the sheet. This value is only available on some sheets where the
   * sheet XML has the dimension data set. At present, this method will return 0 if this
   * dimension data is missing (this may change in a future release).
   *
   * @return first row contained in this sheet (0-based)
   */
  @Override
  public int getFirstRowNum() {
    return reader.getFirstRowNum();
  }

  /**
   * Gets the last row on the sheet. This value is only available on some sheets where the
   * sheet XML has the dimension data set. At present, this method will return 0 if this
   * dimension data is missing (this may change in a future release).
   *
   * @return last row contained in this sheet (0-based)
   */
  @Override
  public int getLastRowNum() {
    return reader.getLastRowNum();
  }

  /**
   * Return cell comment at row, column, if one exists. Otherwise, return null.
   *
   * @param cellAddress the location of the cell comment
   * @return the cell comment, if one exists. Otherwise, return null.
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadComments(boolean)} is not set to true
   */
  @Override
  public Comment getCellComment(CellAddress cellAddress) {
    Comments sheetComments = reader.getCellComments();
    if (sheetComments == null) {
      return null;
    }
    XSSFComment xssfComment = sheetComments.findCellComment(cellAddress);
    if (xssfComment != null && reader.getBuilder().adjustLegacyComments()) {
      return new WrappedComment(xssfComment);
    }
    return xssfComment;
  }

  /**
   * Returns all cell comments on this sheet.
   * @return A map of each Comment in the sheet, keyed on the cell address where
   * the comment is located.
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadComments(boolean)} is not set to true
   */
  @Override
  public Map<CellAddress, ? extends Comment> getCellComments() {
    Comments sheetComments = reader.getCellComments();
    if (sheetComments == null) {
      return Collections.emptyMap();
    }
    Map<CellAddress, Comment> map = new HashMap<>();
    for(Iterator<CellAddress> iter = sheetComments.getCellAddresses(); iter.hasNext(); ) {
      CellAddress address = iter.next();
      map.put(address, getCellComment(address));
    }
    return map;
  }

  /**
   * Only works after sheet is fully read (because merged regions data is stored
   * at the end of the sheet XML).
   */
  @Override
  public CellRangeAddress getMergedRegion(int index) {
    List<CellRangeAddress> regions = getMergedRegions();
    if(index > regions.size()) {
      throw new NoSuchElementException("index " + index + " is out of range");
    }
    return regions.get(index);
  }

  /**
   * Only works after sheet is fully read (because merged regions data is stored
   * at the end of the sheet XML).
   */
  @Override
  public List<CellRangeAddress> getMergedRegions() {
    return reader.getMergedCells();
  }

  /**
   * Only works after sheet is fully read (because merged regions data is stored
   * at the end of the sheet XML).
   */
  @Override
  public int getNumMergedRegions() {
    List<CellRangeAddress> mergedCells = reader.getMergedCells();
    return mergedCells == null ? 0 : mergedCells.size();
  }

  /**
   * Return the sheet's existing drawing, or null if there isn't yet one.
   *
   * @return a SpreadsheetML drawing
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadShapes(boolean)} is not set to true
   */
  @Override
  public Drawing<?> getDrawingPatriarch() {
    return reader.getDrawingPatriarch();
  }

  /**
   * Get a Hyperlink in this sheet anchored at row, column (only if feature is enabled on the Builder).
   *
   * @param row The row where the hyperlink is anchored
   * @param column The column where the hyperlink is anchored
   * @return hyperlink if there is a hyperlink anchored at row, column; otherwise returns null
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  @Override
  public Hyperlink getHyperlink(int row, int column) {
    return getHyperlink(new CellAddress(row, column));
  }

  /**
   * Get hyperlink associated with cell (only if feature is enabled on the Builder).
   * This should only be called after all the rows are read because the hyperlink data is
   * at the end of the sheet.
   *
   * @param cellAddress
   * @return the hyperlink associated with this cell (only if feature is enabled on the Builder) - null if not found
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  @Override
  public Hyperlink getHyperlink(CellAddress cellAddress) {
    for (Hyperlink hyperlink : getHyperlinkList()) {
      if (cellAddress.getRow() >= hyperlink.getFirstRow() && cellAddress.getRow() <= hyperlink.getLastRow()
        && cellAddress.getColumn() >= hyperlink.getFirstColumn() && cellAddress.getColumn() <= hyperlink.getLastColumn()) {
        return hyperlink;
      }
    }
    return null;
  }

  /**
   * Get hyperlinks associated with sheet (only if feature is enabled on the Builder).
   * This should only be called after all the rows are read because the hyperlink data is
   * at the end of the sheet.
   *
   * @return the hyperlinks associated with this sheet (only if feature is enabled on the Builder) - cast to {@link XlsxHyperlink} to access cell reference
   * @throws IllegalStateException if {@link com.github.pjfanning.xlsx.StreamingReader.Builder#setReadHyperlinks(boolean)} is not set to true
   */
  @Override
  public List<? extends Hyperlink> getHyperlinkList() {
    return reader.getHyperlinks();
  }

  @Override
  public CellAddress getActiveCell() {
    return reader.getActiveCell();
  }

  /**
   * Get the default column width for the sheet (if the columns do not define their own width) in
   * characters.
   * <p>
   * Note, this value is different from {@link #getColumnWidth(int)}. The latter is always greater and includes
   * 4 pixels of margin padding (two on each side), plus 1 pixel padding for the gridlines.
   * </p>
   * <p>
   * This value is only available after the first row is read (due to the way excel-streaming-reader streams the data).
   * </p>
   * @return column width, default value is 8
   */
  @Override
  public int getDefaultColumnWidth() {
    return reader.getBaseColWidth();
  }

  /**
   * Get the default row height for the sheet (if the rows do not define their own height) in
   * twips (1/20 of a point)
   * <p>
   * This value is only available after the first row is read (due to the way excel-streaming-reader streams the data).
   * </p>
   *
   * @return  default row height
   */
  @Override
  public short getDefaultRowHeight() {
    return (short)(getDefaultRowHeightInPoints() * Font.TWIPS_PER_POINT);
  }

  /**
   * Get the default row height for the sheet measured in point size (if the rows do not define their own height).
   * <p>
   * This value is only available after the first row is read (due to the way excel-streaming-reader streams the data).
   * </p>
   *
   * @return  default row height in points
   */
  @Override
  public float getDefaultRowHeightInPoints() {
    return reader.getDefaultRowHeight();
  }

  @Override
  public int getColumnWidth(final int columnIndex) {
    return Math.round(reader.getColumnWidth(columnIndex)*256);
  }

  @Override
  public float getColumnWidthInPixels(final int columnIndex) {
    float widthIn256 = getColumnWidth(columnIndex);
    return (float)(widthIn256/256.0 * Units.DEFAULT_CHARACTER_WIDTH);
  }

  @Override
  public PaneInformation getPaneInformation() {
    return reader.getPane();
  }

  /**
   * @return immutable copy of the shared formula map for this sheet
   */
  public Map<String, SharedFormula> getSharedFormulaMap() {
    return reader.getSharedFormulaMap();
  }

  /**
   * @param siValue the ID for the shared formula (appears in Excel sheet XML as an <code>si</code> attribute
   * @param sharedFormula maps the base cell and formula for the shared formula
   */
  public void addSharedFormula(String siValue, SharedFormula sharedFormula) {
    reader.addSharedFormula(siValue, sharedFormula);
  }

  /**
   * @param siValue the ID for the shared formula (appears in Excel sheet XML as an <code>si</code> attribute
   * @return the shared formula that was removed (can be null if no existing shared formula is found)
   */
  public SharedFormula removeSharedFormula(String siValue) {
    return reader.removeSharedFormula(siValue);
  }

}
