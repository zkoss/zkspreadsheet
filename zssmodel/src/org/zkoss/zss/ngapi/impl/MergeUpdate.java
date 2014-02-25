/* MergeChange.java

	Purpose:
		
	Description:
		
	History:
		May 28, 2010 4:35:32 PM, Created by henrichen

Copyright (C) 2010 Potix Corporation. All Rights Reserved.

*/

package org.zkoss.zss.ngapi.impl;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.SSheet;

/**
 * A pair of reference areas indicate the changes of the merge area. 
 * @author henrichen
 * @author dennischen
 */
public class MergeUpdate {
	final private SSheet _sheet;
	final private CellRegion _orgMerge; //original merge range
	final private CellRegion _merge; //merge range changed
	public MergeUpdate(SSheet sheet, CellRegion orgMerge, CellRegion merge) {
		this._sheet = sheet;
		this._orgMerge = orgMerge;
		this._merge = merge;
	}
	public CellRegion getOrgMerge() {
		return _orgMerge;
	}

	public CellRegion getMerge() {
		return _merge;
	}
	
	public SSheet getSheet(){
		return _sheet;
	}
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((_merge == null) ? 0 : _merge.hashCode());
		result = prime * result
				+ ((_orgMerge == null) ? 0 : _orgMerge.hashCode());
		result = prime * result + ((_sheet == null) ? 0 : _sheet.hashCode());
		return result;
	}
	@Override
	public boolean equals(Object obj) {
		if (this == obj)
			return true;
		if (obj == null)
			return false;
		if (getClass() != obj.getClass())
			return false;
		MergeUpdate other = (MergeUpdate) obj;
		if (_merge == null) {
			if (other._merge != null)
				return false;
		} else if (!_merge.equals(other._merge))
			return false;
		if (_orgMerge == null) {
			if (other._orgMerge != null)
				return false;
		} else if (!_orgMerge.equals(other._orgMerge))
			return false;
		if (_sheet == null) {
			if (other._sheet != null)
				return false;
		} else if (!_sheet.equals(other._sheet))
			return false;
		return true;
	}

}
