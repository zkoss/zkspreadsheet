package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

class ColumnProxyImpl extends AbstractColumn{
	private static final long serialVersionUID = 1L;
	private final WeakReference<AbstractSheet> sheetRef;
	private final int index;
	private AbstractColumn proxy;
	
	
	public ColumnProxyImpl(AbstractSheet sheet, int index) {
		this.sheetRef = new WeakReference(sheet);
		this.index = index;
	}
	
	
	protected void loadProxy(){
		if(proxy==null){
			proxy = (AbstractColumn)getSheet().getColumnAt(index,false);
			if(proxy!=null){
				sheetRef.clear();
			}
		}
	}
	
	protected AbstractSheet getSheet(){
		AbstractSheet sheet = sheetRef.get();
		if(sheet==null){
			throw new IllegalStateException("proxy target lost, you should't keep this instance");
		}
		return sheet;
	}
	
	public int getIndex() {
		loadProxy();
		return proxy==null?index:proxy.getIndex();
	}


	public boolean isNull() {
		loadProxy();
		return proxy==null?true:proxy.isNull();
	}
	
	public String asString() {
		loadProxy();
		return proxy==null?CellReference.convertNumToColString(getIndex()):proxy.asString();
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
			proxy = (AbstractColumn)getSheet().getOrCreateColumnAt(index);
		}
		proxy.setCellStyle(cellStyle);
	}
}
