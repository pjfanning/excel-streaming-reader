package com.github.pjfanning.xlsx.impl.adapter;

import org.apache.poi.ss.usermodel.Hyperlink;

public interface HyperlinkAdapter extends Hyperlink {
	/**
	 * @throws UnsupportedOperationException
	 */
	@Override
	default void setAddress(String address) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * @throws UnsupportedOperationException
	 */
	@Override
	default void setLabel(String label) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * @throws UnsupportedOperationException
	 */
	@Override
	default void setFirstColumn(int col) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * @throws UnsupportedOperationException
	 */
	@Override
	default void setLastColumn(int col) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * @throws UnsupportedOperationException
	 */
	@Override
	default void setFirstRow(int row) {
		throw new UnsupportedOperationException("update operations are not supported");
	}

	/**
	 * @throws UnsupportedOperationException
	 */
	@Override
	default void setLastRow(int row) {
		throw new UnsupportedOperationException("update operations are not supported");
	}
}
