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
	
	String getId();
	
//	NCellStyle getCellStyle();
//	
//	NViewInfo getViewInfo();
	
	//editable
	void clearRow(int rowIdx, int rowIdx2);
	void clearColumn(int columnIdx,int columnIdx2);
	void clearCell(int rowIdx, int columnIdx,int rowIdx2,int columnIdx2);
	
	void insertRow(int rowIdx, int size);
	void deleteRow(int rowIdx, int size);
	void insertColumn(int columnIdx, int size);
	void deleteColumn(int columnIdx, int size);
	
	
}
