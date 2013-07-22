/* ChartImpl.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/5/1 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.api.model.impl;

import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.impl.RangeImpl;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XSheet;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class ChartImpl implements Chart{
	
	ModelRef<XSheet> sheetRef;
	ModelRef<org.zkoss.poi.ss.usermodel.Chart> chartRef;
	
	public ChartImpl(ModelRef<XSheet> sheetRef, ModelRef<org.zkoss.poi.ss.usermodel.Chart> chartRef) {
		this.sheetRef = sheetRef;
		this.chartRef = chartRef;
	}

	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((chartRef == null) ? 0 : chartRef.hashCode());
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
		ChartImpl other = (ChartImpl) obj;
		if (chartRef == null) {
			if (other.chartRef != null)
				return false;
		} else if (!chartRef.equals(other.chartRef))
			return false;
		return true;
	}
	
	public org.zkoss.poi.ss.usermodel.Chart getNative() {
		return chartRef.get();
	}
	
	public String getId(){
		return getNative().getChartId();
	}
	
	@Override
	public SheetAnchor getAnchor() {
		ClientAnchor anchor = getNative().getPreferredSize();
		return anchor==null?null:SheetImpl.toSheetAnchor(sheetRef.get(), anchor);
	}
}
