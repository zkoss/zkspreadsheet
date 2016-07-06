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
package org.zkoss.zss.model.impl;

import java.io.Serializable;
import java.util.Collection;
import java.util.List;
import java.util.Set;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SColumn;
import org.zkoss.zss.model.SConditionalFormattingRule;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.STable;
import org.zkoss.zss.model.SConditionalFormatting;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public abstract class AbstractSheetAdv implements SSheet,LinkedModelObject,Serializable{
	private static final long serialVersionUID = 1L;
	
	/*package*/ abstract AbstractRowAdv getRow(int rowIdx, boolean proxy);
	/*package*/ abstract AbstractRowAdv getOrCreateRow(int rowIdx);
//	/*package*/ abstract int getRowIndex(AbstractRowAdv row);
	
	/*package*/ abstract SColumn getColumn(int columnIdx, boolean proxy);
	/*package*/ abstract AbstractColumnArrayAdv getOrSplitColumnArray(int index);
	
//	/*package*/ abstract ColumnAdv getOrCreateColumn(int columnIdx);
//	/*package*/ abstract int getColumnIndex(ColumnAdv column);
	
	/*package*/ abstract AbstractCellAdv getCell(int rowIdx, int columnIdx, boolean proxy);
	/*package*/ abstract AbstractCellAdv getOrCreateCell(int rowIdx, int columnIdx);
	
	
	/*package*/ abstract void copyTo(AbstractSheetAdv sheet);
	/*package*/ abstract void setSheetName(String name);
	
//	/*package*/ abstract void onModelInternalEvent(ModelInternalEvent event);
	
	//ZSS-855
	abstract public STable getTableByRowCol(int rowIdx, int colIdx);
	
	//ZSS-962
	abstract public boolean isHidden(int rowIdx, int colIdx);
	
	//ZSS-985
	abstract public void removeTables(Set<String> tableNames);
	
	//ZSS-1001
	abstract public void removeTable(STable table);
	
	//ZSS-1001
	abstract public void clearTables();
	
	//ZSS-1130
	//@since 3.8.2
	abstract public void addConditionalFormatting(SConditionalFormatting scf);
	
	//ZSS-1168
	//@since 3.8.3
	abstract public void setMergeOutOfSync(int state);
	
	//ZSS-1168
	//@since 3.8.3
	abstract public int getMergeOutOfSync();
	
	//ZSS-1142
	//@since 3.9.0
	//@Internal
	abstract public ConditionalStyleImpl getConditionalFormattingStyle(int row, int col);

	//ZSS-1142
	/**
	 * @param row
	 * @param column
	 * @return the associated conditionalFormattingRule
	 * @since 3.9.0
	 */
	abstract public SConditionalFormatting getConditionalFormatting(int row, int column);
	
	//ZSS-1251
	//@since 3.9.0
	//@Internal
	abstract public void removeConditionalFormatting(SConditionalFormatting scf);

	//ZSS-1251
	//@since 3.9.0
	//@Internal
	abstract public int nextConditionalId();
	
	//ZSS-1251
	//@since 3.9.0
	//@Internal
	abstract public SConditionalFormatting getConditionalFormatting(int id);

	//ZSS-1251
	/**
	 * Delete a conditional formatting from this sheet.
	 * @param cfmt
	 * @since 3.9.0
	 */
	abstract public void deleteConditionalFormatting(SConditionalFormatting cfmt);

	//ZSS-1251
	/**
	 * Remove a region from conditional formatting.
	 * @param region
	 * @since 3.9.0
	 */
	abstract public void removeConditionalFormattingRegion(CellRegion region);

	//ZSS-1251
	/**
	 * Delete a region from conditional formatting and return the deleted 
	 * conditional formatting.
	 * @param region
	 * @return
	 * @since 3.9.0
	 */
	abstract public List<SConditionalFormatting> deleteConditionalFormattingRegion(CellRegion region);
	
	//ZSS-1251
	/**
	 * Paste from src a new ConditionalFormatting at the specified region.
	 * @param region
	 * @param src
	 * @return
	 * @since 3.9.0
	 */
	abstract public SConditionalFormatting addConditionalFormatting(CellRegion srcrgn, CellRegion dstrgn, SConditionalFormatting src, int rowOff, int colOff);
}
