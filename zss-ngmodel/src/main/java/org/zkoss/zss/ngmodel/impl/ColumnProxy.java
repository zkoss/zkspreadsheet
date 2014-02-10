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

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumn;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
class ColumnProxy implements NColumn {
	private static final long serialVersionUID = 1L;
	private final WeakReference<AbstractSheetAdv> sheetRef;
	private final int index;
	private AbstractColumnArrayAdv proxy;

	public ColumnProxy(AbstractSheetAdv sheet, int index) {
		this.sheetRef = new WeakReference(sheet);
		this.index = index;
	}

	protected void loadProxy(boolean split) {
		if(split){
			proxy = (AbstractColumnArrayAdv) ((AbstractSheetAdv)getSheet()).getOrSplitColumnArray(index);
		}else if (proxy == null) {
			proxy = (AbstractColumnArrayAdv) ((AbstractSheetAdv)getSheet()).getColumnArray(index);
		}
	}

	@Override
	public NSheet getSheet() {
		AbstractSheetAdv sheet = sheetRef.get();
		if (sheet == null) {
			throw new IllegalStateException(
					"proxy target lost, you should't keep this instance");
		}
		return sheet;
	}

	@Override
	public int getIndex() {
		return index;
	}

	@Override
	public boolean isNull() {
		loadProxy(false);
		return proxy == null ? true : false;
	}

	@Override
	public NCellStyle getCellStyle() {
		loadProxy(false);
		if (proxy != null) {
			return proxy.getCellStyle();
		}
		return getSheet().getBook().getDefaultCellStyle();
	}

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		loadProxy(true);
		proxy.setCellStyle(cellStyle);
	}

	@Override
	public int getWidth() {
		loadProxy(false);
		if (proxy != null) {
			return proxy.getWidth();
		}
		return getSheet().getDefaultColumnWidth();
	}

	@Override
	public boolean isHidden() {
		loadProxy(false);
		if (proxy != null) {
			return proxy.isHidden();
		}
		return false;
	}

	@Override
	public void setWidth(int width) {
		loadProxy(true);
		proxy.setWidth(width);
	}

	@Override
	public void setHidden(boolean hidden) {
		loadProxy(true);
		proxy.setHidden(hidden);
	}
	
	@Override
	public boolean isCustomWidth() {
		loadProxy(false);
		if (proxy != null) {
			return proxy.isCustomWidth();
		}
		return false;
	}
	@Override
	public void setCustomWidth(boolean custom) {
		loadProxy(true);
		proxy.setCustomWidth(custom);
	}
}
