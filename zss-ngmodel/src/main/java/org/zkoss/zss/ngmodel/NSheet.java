package org.zkoss.zss.ngmodel;

/**
 * @author dennis
 *
 */
public interface NSheet {
	
	NBook getBook();
	
	String getSheetName();
	
	NRow getRow(int rowIdx);
	
	NColumn getColumn(int columnIdx);
	
	NCell getCell(int rowIdx, int columnIdx);
	
	int getStartRowIndex();
	int getEndRowIndex();
	int getStartColumnIndex();
	int getEndColumnIndex();
	
	int getStartColumnIndex(int rowIdx);
	int getEndColumn(int rowIdx);
	
	
//	NCellStyle getCellStyle();
//	
//	NViewInfo getViewInfo();
	
	//editable
	void clearRow(int rowIdx, int rowIdx2);
	void clearColumn(int columnIdx,int columnIdx2);
	void clearCell(int rowIdx, int columnIdx,int rowIdx2,int columnIdx2);
	
//	void insertRow(int rowFrom, int size);
//	void deleteRow(int rowFrom, int size);
//	void insertColumn(int columnFrom, int size);
//	void deleteColumn(int columnFrom, int size);
	
	
}
