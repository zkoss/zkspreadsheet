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
package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class ColumnArrayImpl extends AbstractColumnArrayAdv {
	private static final long serialVersionUID = 1L;

	private AbstractSheetAdv sheet;
	private AbstractCellStyleAdv cellStyle;
	
	private Integer width;
	private boolean hidden = false;
	private boolean customWidth = false;

	int index;
	int lastIndex;
	
	public ColumnArrayImpl(AbstractSheetAdv sheet, int index, int lastIndex) {
		this.sheet = sheet;
		this.index = index;
		this.lastIndex = lastIndex;
	}

	@Override
	public int getIndex() {
		checkOrphan();
		return index;
	}


//	@Override
//	void onModelInternalEvent(ModelInternalEvent event) {
//		// TODO Auto-generated method stub
//
//	}


	@Override
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	public void destroy() {
		checkOrphan();
		sheet = null;
	}

	@Override
	public NSheet getSheet() {
		checkOrphan();
		return sheet;
	}

	@Override
	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	@Override
	public NCellStyle getCellStyle(boolean local) {
		if (local || cellStyle != null) {
			return cellStyle;
		}
		checkOrphan();
		return sheet.getBook().getDefaultCellStyle();
	}

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, AbstractCellStyleAdv.class);
		this.cellStyle = (AbstractCellStyleAdv) cellStyle;
	}

	@Override
	public int getWidth() {
		if(width!=null){
			return width.intValue();
		}
		checkOrphan();
		return getSheet().getDefaultColumnWidth();
	}

	@Override
	public boolean isHidden() {
		return hidden;
	}

	@Override
	public void setWidth(int width) {
		this.width = Integer.valueOf(width);
	}

	@Override
	public void setHidden(boolean hidden) {
		this.hidden = hidden;
	}

	@Override
	public int getLastIndex() {
		return lastIndex;
	}

	@Override
	void setIndex(int index) {
		this.index = index;
	}

	@Override
	void setLastIndex(int lastIndex) {
		this.lastIndex = lastIndex;
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append(super.toString()).append("[").append(getIndex()).append("-").append(getLastIndex()).append("]");
		return sb.toString();
	}

	@Override
	public boolean isCustomWidth() {
		return customWidth;
	}

	@Override
	public void setCustomWidth(boolean custom) {
		customWidth = custom;
	}

}
