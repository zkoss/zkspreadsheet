package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.util.CellReference;

public class RefImpl implements Ref, Serializable {

	private static final long serialVersionUID = 1L;
	RefType type;
	String bookName;
	String sheetName;
	int row = -1;
	int column = -1;
	int lastRow = -1;
	int lastColumn = -1;
	String objectId;

	public RefImpl(String bookName, String sheetName, int row, int column,
			int lastRow, int lastColumn) {
		setInfo(RefType.AREA, bookName, sheetName, row, column, lastRow,
				lastColumn, null);
	}

	public RefImpl(String bookName, String sheetName, int row, int column) {
		setInfo(RefType.CELL, bookName, sheetName, row, column, row, column,
				null);
	}

	public RefImpl(String bookName, String sheetName) {
		setInfo(RefType.SHEET, bookName, sheetName, -1, -1, -1, -1, null);
	}

	public RefImpl(String bookName) {
		setInfo(RefType.BOOK, bookName, null, -1, -1, -1, -1, null);
	}

	public RefImpl(AbstractCell cell) {
		AbstractSheet sheet = (AbstractSheet)cell.getSheet();
		AbstractBook book = ((AbstractBook) sheet.getBook());
		int row = cell.getRowIndex();
		int column = cell.getColumnIndex();
		setInfo(RefType.CELL, book.getBookName(), sheet.getSheetName(), row,
				column, row, column, null);
	}

	public RefImpl(AbstractSheet sheet) {
		AbstractBook book = ((AbstractBook) sheet.getBook());
		setInfo(RefType.SHEET, book.getBookName(), sheet.getSheetName(), -1,
				-1, -1, -1, null);
	}

	public RefImpl(AbstractBook book) {
		setInfo(RefType.BOOK, book.getBookName(), null, -1, -1, -1, -1, null);
	}

	public RefImpl(AbstractChart chart) {
		setInfo(RefType.CHART, chart.getSheet().getBook().getBookName(), null,
				-1, -1, -1, -1, chart.getId());
	}

	private void setInfo(RefType type, String bookName, String sheetName,
			int row, int column, int lastRow, int lastColumn, String objectId) {
		this.type = type;
		this.bookName = bookName;
		this.sheetName = sheetName;
		this.row = row;
		this.column = column;
		this.lastRow = lastRow;
		this.lastColumn = lastColumn;
		this.objectId = objectId;
	}

	@Override
	public RefType getType() {
		return type;
	}

	@Override
	public String getBookName() {
		return bookName;
	}

	@Override
	public String getSheetName() {
		return sheetName;
	}

	@Override
	public int getRow() {
		return row;
	}

	@Override
	public int getColumn() {
		return column;
	}

	@Override
	public int getLastRow() {
		return lastRow;
	}

	@Override
	public int getLastColumn() {
		return lastColumn;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result
				+ ((bookName == null) ? 0 : bookName.hashCode());
		result = prime * result + column;
		result = prime * result + lastColumn;
		result = prime * result + lastRow;
		result = prime * result + row;
		result = prime * result
				+ ((sheetName == null) ? 0 : sheetName.hashCode());
		result = prime * result + ((type == null) ? 0 : type.hashCode());
		return result;
	}

	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		RefImpl other = (RefImpl) obj;
		if (bookName == null) {
			if (other.bookName != null)
				return false;
		} else if (!bookName.equals(other.bookName))
			return false;
		if (column != other.column)
			return false;
		if (lastColumn != other.lastColumn)
			return false;
		if (lastRow != other.lastRow)
			return false;
		if (row != other.row)
			return false;
		if (sheetName == null) {
			if (other.sheetName != null)
				return false;
		} else if (!sheetName.equals(other.sheetName))
			return false;
		if (type != other.type)
			return false;
		return true;
	}

	@Override
	public String toString() {
		StringBuilder sb = new StringBuilder();
		switch (type) {
		case AREA:
			sb.insert(
					0,
					":"
							+ new CellReference(lastRow, lastColumn)
									.formatAsString());
		case CELL:
			sb.insert(0, new CellReference(row, column).formatAsString());
		case SHEET:
			sb.insert(0, sheetName + "!");
			break;
		case CHART:
		case VALIDATION:
			sb.insert(0, objectId);
		}

		sb.insert(0, bookName + ":");
		return sb.toString();
	}

	@Override
	public String getObjectId() {
		return objectId;
	}

}
