/* SConditionalFormatting.java

	Purpose:
		
	Description:
		
	History:
		Oct 23, 2015 11:51:24 AM, Created by henrichen

	Copyright (C) 2015 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model;

import java.util.List;
import java.util.Set;

/**
 * Conditional Formatting
 * @author henri
 *
 */
public interface SConditionalFormatting {
	/**
	 * The sheet on which this conditional formatting covered
	 * @return
	 */
	SSheet getSheet();
	
	/**
	 * Regions that this conditional formatting covered
	 * @return
	 */
	Set<CellRegion> getRegions();
	
	/**
	 * Rules applied to the covered region
	 * @return
	 */
	List<SConditionalFormattingRule> getRules();
	
	/**
	 * Remove partial region in the ConditionalFormatting
	 * @param region
	 * @since 3.9.0
	 */
	void removeRegion(CellRegion region);
	
	/**
	 * When sheet name changed; must call this to update region and rule formula 
	 * @param oldName
	 * @param newName
	 * @since 3.9.0
	 */
	void renameSheet(String oldName, String newName);

	/**
	 * Call this destory() when this formatting is about to be destroied.
	 * @since 3.9.0
	 */
	public void destroy();
	
	/**
	 * Used to copy the contents from another src ConditionalFormatting. 
	 * @param src
	 * @since 3.9.0
	 */
	void copyFrom(SConditionalFormatting src, int rowOff, int colOff);
	
	/**
	 * 
	 * @param regions
	 * @since 3.9.0
	 */
	void setRegions(Set<CellRegion> regions);
	
	/**
	 * 
	 * @param region
	 * @since 3.9.0
	 */
	void addRegion(CellRegion region);
	
	/**
	 * 
	 * @return
	 * @since 3.9.0
	 */
	public int getId();

	/**
	 * @since 3.9.0
	 */
	public void clearFormulaResultCache();

}
