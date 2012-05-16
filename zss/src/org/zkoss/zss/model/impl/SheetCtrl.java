/* SheetCtrl.java

	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 11:10:50 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import org.zkoss.poi.ss.usermodel.PivotTable;
import org.zkoss.poi.ss.util.CellRangeAddress;

/**
 * Sheet controls (Internal Use only).
 * @author henrichen
 *
 */
public interface SheetCtrl {
	/**
	 * Evaluate all formulas of the associated sheet.
	 */
	public void evalAll();
	
	/**
	 * Returns whether the associated sheet has evaluated all formulas.
	 * @return whether the associated sheet has evaluated all formulas. 
	 */
	public boolean isEvalAll();
    
    /**
     * Returns the universal unique id of the associated Sheet.
     * @return the universal unique id of the associated Sheet.
     */
    public String getUuid();
    
    /**
     * Initialize the merge cells.
     */
    public void initMerged();
    
    /*
     * Returns merged region of the given row, column as the left top cell of the merged region; if not a merged cell, return null.
     * @param row row index
     * @param col column index
     * @return merged region of the given row, column as the left top cell of the merged region; if not a merged cell, return null.
     */
    public CellRangeAddress getMerged(int row, int col);
    
    /**
     * Add a new merged region.
     * @param addr merged region
     */
	public void addMerged(CellRangeAddress addr);
	
    /**
     * Delete a merged region.
     * @param addr to be removed merged region
     */
	public void deleteMerged(CellRangeAddress addr);
	
	/**
	 * Return associated drawing manager.
	 * @return drawing manager
	 */
	public DrawingManager getDrawingManager();
	
	public PivotTableManager getPivotTableManager();
	
	/**
	 * Callback when the name of a sheet in the associated book changes.
	 * @param oldname old sheet name
	 * @param newname new sheet name
	 */
	public void whenRenameSheet(String oldname, String newname);
}
