package org.zkoss.zss.ngmodel.sys.dependency;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface Ref {

	public enum RefType {
		CELL, AREA, SHEET, BOOK, OBJECT
	}

	public RefType getType();

	public String getBookName();

	public String getSheetName();

	public int getRow();

	public int getColumn();

	public int getLastRow();

	public int getLastColumn();
}
