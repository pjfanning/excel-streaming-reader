package com.github.pjfanning.xlsx.impl.adapter;

import com.github.pjfanning.xlsx.exceptions.NotSupportedException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;

public interface RowAdapter extends Row {

	/**
	 * Not supported
	 */
	@Override
	default Cell createCell(int column) {
		throw new NotSupportedException();
	}

	/**
	 * Not supported
	 */
	@Override
	default Cell createCell(int i, CellType cellType) {
		throw new NotSupportedException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void removeCell(Cell cell) {
		throw new NotSupportedException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setRowNum(int rowNum) {
		throw new NotSupportedException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setHeight(short height) {
		throw new NotSupportedException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setZeroHeight(boolean zHeight) {
		throw new NotSupportedException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void setHeightInPoints(float height) {
		throw new NotSupportedException("update operations are not supported");
	}

	/**
	 * Not supported
	 */
	@Override
	default int getOutlineLevel() {
		throw new NotSupportedException();
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void shiftCellsRight(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
		throw new NotSupportedException("update operations are not supported");
	}

	/**
	 * Update operations are not supported
	 */
	@Override
	default void shiftCellsLeft(int firstShiftColumnIndex, int lastShiftColumnIndex, int step) {
		throw new NotSupportedException("update operations are not supported");
	}
}
