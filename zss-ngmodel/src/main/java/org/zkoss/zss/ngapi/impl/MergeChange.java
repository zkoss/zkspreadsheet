/* MergeChange.java

	Purpose:
		
	Description:
		
	History:
		May 28, 2010 4:35:32 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.ngmodel.CellRegion;

/**
 * A pair of reference areas indicate the changes of the merge area. 
 * @author henrichen
 *
 */
public class MergeChange {
	final private CellRegion _orgMerge; //original merge range
	final private CellRegion _merge; //merge range changed
	public MergeChange(CellRegion orgMerge, CellRegion merge) {
		this._orgMerge = orgMerge;
		this._merge = merge;
	}
	public CellRegion getOrgMerge() {
		return _orgMerge;
	}

	public CellRegion getMerge() {
		return _merge;
	}
}
