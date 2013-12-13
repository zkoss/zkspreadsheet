/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;

import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.util.CellReference;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class RefImpl implements Ref, Serializable {

	private static final long serialVersionUID = 1L;
	final private RefType type;
	final protected String bookName;
	final protected String sheetName;
	final private int row;
	final private int column;
	final private int lastRow;
	final private int lastColumn;

	public RefImpl(String bookName, String sheetName, int row, int column,
			int lastRow, int lastColumn) {
		this(RefType.AREA, bookName, sheetName, row, column, lastRow,lastColumn);
	}

	public RefImpl(String bookName, String sheetName, int row, int column) {
		this(RefType.CELL, bookName, sheetName, row, column, row, column);
	}

	public RefImpl(String bookName, String sheetName) {
		this(RefType.SHEET, bookName, sheetName, -1, -1, -1, -1);
	}

	public RefImpl(String bookName) {
		this(RefType.BOOK, bookName, null, -1, -1, -1, -1);
	}

	public RefImpl(CellAdv cell) {
		this(RefType.CELL, cell.getSheet().getBook().getBookName(), cell.getSheet().getSheetName(), cell.getRowIndex(),
		cell.getColumnIndex(), cell.getRowIndex(), cell.getColumnIndex());
	}

	public RefImpl(SheetAdv sheet) {
		this(RefType.SHEET, ((BookAdv) sheet.getBook()).getBookName(), sheet.getSheetName(), -1, -1, -1, -1);
	}

	public RefImpl(BookAdv book) {
		this(RefType.BOOK, book.getBookName(), null, -1, -1, -1, -1);
	}

	protected RefImpl(RefType type, String bookName, String sheetName,
			int row, int column, int lastRow, int lastColumn) {
		this.type = type;
		this.bookName = bookName;
		this.sheetName = sheetName;
		this.row = row;
		this.column = column;
		this.lastRow = lastRow;
		this.lastColumn = lastColumn;
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
			sb.insert(0,":"+ new CellReference(lastRow, lastColumn).formatAsString());
		case CELL:
			sb.insert(0, new CellReference(row, column).formatAsString());
		case SHEET:
			sb.insert(0, sheetName + "!");
			break;
		case OBJECT://will be override
		case NAME://will be override
		case BOOK:
		}

		sb.insert(0, bookName + ":");
		return sb.toString();
	}
}
