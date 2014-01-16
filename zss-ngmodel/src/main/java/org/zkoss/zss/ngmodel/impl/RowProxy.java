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
import java.util.Collections;
import java.util.Iterator;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
class RowProxy extends AbstractRowAdv{
	private static final long serialVersionUID = 1L;
	
	private final WeakReference<AbstractSheetAdv> sheetRef;
	private final int index;
	AbstractRowAdv proxy;
	
	public RowProxy(AbstractSheetAdv sheet, int index) {
		this.sheetRef = new WeakReference(sheet);
		this.index = index;
	}
	@Override
	public NSheet getSheet(){
		AbstractSheetAdv sheet = sheetRef.get();
		if(sheet==null){
			throw new IllegalStateException("proxy target lost, you should't keep this instance");
		}
		return sheet;
	}
	
	protected void loadProxy(){
		if(proxy==null){
			proxy = (AbstractRowAdv)((AbstractSheetAdv)getSheet()).getRow(index,false);
			if(proxy!=null){
				sheetRef.clear();
			}
		}
	}
	
	public int getIndex() {
		loadProxy();
		return proxy==null?index:proxy.getIndex();
	}


	public boolean isNull() {
		loadProxy();
		return proxy==null?true:proxy.isNull();
	}


	public int getStartCellIndex() {
		loadProxy();
		return proxy==null?-1:proxy.getStartCellIndex();
	}


	public int getEndCellIndex() {
		loadProxy();
		return proxy==null?-1:proxy.getEndCellIndex();
	}
	
	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	public NCellStyle getCellStyle(boolean local) {
		loadProxy();
		if(proxy!=null){
			return proxy.getCellStyle(local);
		}
		return local?null:getSheet().getBook().getDefaultCellStyle();
	}
	
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		loadProxy();
		if(proxy==null){
			proxy = (AbstractRowAdv)((AbstractSheetAdv)getSheet()).getOrCreateRow(index);
		}
		proxy.setCellStyle(cellStyle);
	}
	
	@Override
	AbstractCellAdv getCell(int columnIdx, boolean proxy) {
		throw new UnsupportedOperationException("readonly");
	}
	@Override
	AbstractCellAdv getOrCreateCell(int columnIdx) {
		throw new UnsupportedOperationException("readonly");
	}
	@Override
	void clearCell(int start, int end) {
		throw new UnsupportedOperationException("readonly");
	}
	@Override
	void insertCell(int start, int size) {
		throw new UnsupportedOperationException("readonly");
	}
	@Override
	void deleteCell(int start, int size) {
		throw new UnsupportedOperationException("readonly");
	}
	@Override
	public void destroy() {
		throw new IllegalStateException("never link proxy object and call it's release");
	}
	@Override
	public void checkOrphan() {}
	
	@Override
	void onModelInternalEvent(ModelInternalEvent event) {}
	
	@Override
	public int getHeight() {
		loadProxy();
		if (proxy != null) {
			return proxy.getHeight();
		}
		return getSheet().getDefaultRowHeight();
	}

	@Override
	public boolean isHidden() {
		loadProxy();
		if (proxy != null) {
			return proxy.isHidden();
		}
		return false;
	}

	@Override
	public void setHeight(int width) {
		loadProxy();
		if (proxy == null) {
			proxy = (AbstractRowAdv)((AbstractSheetAdv)getSheet()).getOrCreateRow(index);
		}
		proxy.setHeight(width);
	}

	@Override
	public void setHidden(boolean hidden) {
		loadProxy();
		if (proxy == null) {
			proxy = (AbstractRowAdv)((AbstractSheetAdv)getSheet()).getOrCreateRow(index);
		}
		proxy.setHidden(hidden);
	}
	
	@Override
	public boolean isCustomHeight() {
		loadProxy();
		if (proxy != null) {
			return proxy.isCustomHeight();
		}
		return false;
	}
	@Override
	public void setCustomHeight(boolean custom) {
		loadProxy();
		proxy.setCustomHeight(custom);
	}
	@Override
	public Iterator<AbstractCellAdv> getCellIterator(boolean reverse) {
		loadProxy();
		if (proxy != null) {
			return proxy.getCellIterator(reverse);
		}
		return Collections.EMPTY_LIST.iterator();
	}
	@Override
	void setIndex(int newidx) {
		throw new UnsupportedOperationException("readonly");
	}
	@Override
	void moveCellTo(AbstractRowAdv target, int start, int end, int offset) {
		throw new UnsupportedOperationException("readonly");
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("Row:").append(getIndex());
		return sb.toString();
	}
}
