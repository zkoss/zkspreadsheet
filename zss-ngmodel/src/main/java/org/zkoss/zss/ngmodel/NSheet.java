package org.zkoss.zss.ngmodel;

/**
 * @author dennis
 *
 */
public interface NSheet {
	
	NBook getBook();
	
	String getSheetName();
	
	NRow getRowAt(int row);
	
	NColumn getColumnAt(int column);
	
	NCell getCellAt(int row, int column);
	
	int getStartRow();
	int getEndRow();
	int getStartColumn();
	int getEndColumn();
	
	int getStartColumn(int row);
	int getEndColumn(int row);
	
	
//	NCellStyle getCellStyle();
//	
//	NViewInfo getViewInfo();
	
	//editable
//	void insertRow(int rowFrom, int size);
//	void deleteRow(int rowFrom, int size);
//	void insertColumn(int columnFrom, int size);
//	void deleteColumn(int columnFrom, int size);
	
	
}
