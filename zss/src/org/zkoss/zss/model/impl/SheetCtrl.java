/* SheetCtrl.java

	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 11:10:50 AM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

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
}
