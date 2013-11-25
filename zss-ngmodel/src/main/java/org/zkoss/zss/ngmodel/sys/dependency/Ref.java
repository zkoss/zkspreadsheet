package org.zkoss.zss.ngmodel.sys.dependency;


public interface Ref {

	public enum RefType {
		CELL, AREA, SHEET, BOOK
	}
	public RefType getType();
	public String getBookName();
	public String getSheetName();	
	public int getRow();
	public int getColumn();
	public int getLastRow();
	public int getLastColumn();
}
