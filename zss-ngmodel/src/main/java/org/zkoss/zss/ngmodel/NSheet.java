package org.zkoss.zss.ngmodel;

/**
 * @author dennis
 *
 */
public interface NSheet {
	
	public NBook getBook();
	
	public String getSheetName();
	
	public NRow getRow(int rowIdx);
	
	public NColumn getColumn(int columnIdx);
	
	public NCell getCell(int rowIdx, int columnIdx);
	
	public int getStartRowIndex();
	public int getEndRowIndex();
	public int getStartColumnIndex();
	public int getEndColumnIndex();
	
	public int getStartColumnIndex(int rowIdx);
	public int getEndColumn(int rowIdx);
	
	public String getId();
	
//	NCellStyle getCellStyle();
//	
//	NViewInfo getViewInfo();
	
	//editable
	public void clearRow(int rowIdx, int rowIdx2);
	public void clearColumn(int columnIdx,int columnIdx2);
	public void clearCell(int rowIdx, int columnIdx,int rowIdx2,int columnIdx2);
	
	public void insertRow(int rowIdx, int size);
	public void deleteRow(int rowIdx, int size);
	public void insertColumn(int columnIdx, int size);
	public void deleteColumn(int columnIdx, int size);
	
	
}
