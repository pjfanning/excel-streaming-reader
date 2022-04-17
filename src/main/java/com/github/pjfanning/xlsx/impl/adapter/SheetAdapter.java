package com.github.pjfanning.xlsx.impl.adapter;

import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellAddress;
import org.apache.poi.ss.util.CellRangeAddress;

import java.util.Collection;
import java.util.List;

public interface SheetAdapter extends Sheet {


	/**
	 * Update operations are not supported
	 */
	@Override
	default Row createRow(int rownum) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void removeRow(Row row) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported - use {@link #iterator()} or {@link #rowIterator()} instead
	 */
	@Override
	default Row getRow(int rownum) {
		throw new UnsupportedOperationException("use iterator() or rowIterator() instead");
	}

	/**
	 * Not supported - use {@link #iterator()} or {@link #rowIterator()} to iterate over rows
	 * and count the rows
	 */
	@Override
	default int getPhysicalNumberOfRows() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setColumnHidden(int columnIndex, boolean hidden) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRightToLeft(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isRightToLeft() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setColumnWidth(int columnIndex, int width) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDefaultColumnWidth(int width) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDefaultRowHeight(short height) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDefaultRowHeightInPoints(float height) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default CellStyle getColumnStyle(int column) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default int addMergedRegion(CellRangeAddress region) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default int addMergedRegionUnsafe(CellRangeAddress cellRangeAddress) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default void validateMergedRegions() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setVerticallyCenter(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setHorizontallyCenter(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getHorizontallyCenter() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getVerticallyCenter() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void removeMergedRegion(int index) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void removeMergedRegions(Collection<Integer> collection) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setForceFormulaRecalculation(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getForceFormulaRecalculation() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setAutobreaks(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDisplayGuts(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDisplayZeros(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isDisplayZeros() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setFitToPage(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRowSumsBelow(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRowSumsRight(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getAutobreaks() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getDisplayGuts() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getFitToPage() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getRowSumsBelow() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getRowSumsRight() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isPrintGridlines() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setPrintGridlines(boolean show) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isPrintRowAndColumnHeadings() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setPrintRowAndColumnHeadings(boolean b) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default PrintSetup getPrintSetup() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Header getHeader() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Footer getFooter() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setSelected(boolean value) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default double getMargin(short margin) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setMargin(short margin, double size) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getProtect() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void protectSheet(String password) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean getScenarioProtect() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setZoom(int i) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default short getTopRow() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default short getLeftCol() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void showInPane(int toprow, int leftcol) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void shiftRows(int startRow, int endRow, int n) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void shiftRows(int startRow, int endRow, int n, boolean copyRowHeight, boolean resetOriginalRowHeight) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void shiftColumns(int startColumn, int endColumn, final int n) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default void createFreezePane(int colSplit, int rowSplit, int leftmostColumn, int topRow) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void createFreezePane(int colSplit, int rowSplit) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void createSplitPane(int xSplitPos, int ySplitPos, int leftmostColumn, int topRow, int activePane) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDisplayGridlines(boolean show) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isDisplayGridlines() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDisplayFormulas(boolean show) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isDisplayFormulas() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDisplayRowColHeadings(boolean show) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isDisplayRowColHeadings() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRowBreak(int row) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isRowBroken(int row) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void removeRowBreak(int row) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default int[] getRowBreaks() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default int[] getColumnBreaks() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setColumnBreak(int column) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isColumnBroken(int column) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void removeColumnBreak(int column) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setColumnGroupCollapsed(int columnNumber, boolean collapsed) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void groupColumn(int fromColumn, int toColumn) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void ungroupColumn(int fromColumn, int toColumn) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void groupRow(int fromRow, int toRow) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void ungroupRow(int fromRow, int toRow) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRowGroupCollapsed(int row, boolean collapse) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setDefaultColumnStyle(int column, CellStyle style) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void autoSizeColumn(int column) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void autoSizeColumn(int column, boolean useMergedCells) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default Drawing createDrawingPatriarch() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default boolean isSelected() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default CellRange<? extends Cell> setArrayFormula(String formula, CellRangeAddress range) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default CellRange<? extends Cell> removeArrayFormula(Cell cell) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default DataValidationHelper getDataValidationHelper() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default List<? extends DataValidation> getDataValidations() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void addValidationData(DataValidation dataValidation) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default AutoFilter setAutoFilter(CellRangeAddress range) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default SheetConditionalFormatting getSheetConditionalFormatting() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default CellRangeAddress getRepeatingRows() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default CellRangeAddress getRepeatingColumns() {
		throw new UnsupportedOperationException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRepeatingRows(CellRangeAddress rowRangeRef) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRepeatingColumns(CellRangeAddress columnRangeRef) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default int getColumnOutlineLevel(int columnIndex) {
		throw new UnsupportedOperationException();
	}

	/**
	 * Not supported
	 */
	@Override
	default void setActiveCell(CellAddress cellAddress) {
		throw new UnsupportedOperationException("update operations are not supported");
	}
}
