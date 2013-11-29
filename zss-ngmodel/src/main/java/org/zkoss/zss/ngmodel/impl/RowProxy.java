package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.ModelEvent;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.util.Validations;

class RowProxy extends RowAdv{
	private static final long serialVersionUID = 1L;
	
	private final WeakReference<SheetAdv> sheetRef;
	private final int index;
	RowAdv proxy;
	
	public RowProxy(SheetAdv sheet, int index) {
		this.sheetRef = new WeakReference(sheet);
		this.index = index;
	}
	@Override
	public NSheet getSheet(){
		SheetAdv sheet = sheetRef.get();
		if(sheet==null){
			throw new IllegalStateException("proxy target lost, you should't keep this instance");
		}
		return sheet;
	}
	
	protected void loadProxy(){
		if(proxy==null){
			proxy = (RowAdv)((SheetAdv)getSheet()).getRow(index,false);
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

	public String asString() {
		loadProxy();
		return proxy==null?Integer.toString(index+1):proxy.asString();
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
			proxy = (RowAdv)((SheetAdv)getSheet()).getOrCreateRow(index);
		}
		proxy.setCellStyle(cellStyle);
	}
	
	@Override
	CellAdv getCell(int columnIdx, boolean proxy) {
		throw new UnsupportedOperationException("not implement");
	}
	@Override
	CellAdv getOrCreateCell(int columnIdx) {
		throw new UnsupportedOperationException("not implement");
	}
	@Override
	void clearCell(int start, int end) {
		throw new UnsupportedOperationException("not implement");
	}
	@Override
	void insertCell(int start, int size) {
		throw new UnsupportedOperationException("not implement");
	}
	@Override
	void deleteCell(int start, int size) {
		throw new UnsupportedOperationException("not implement");
	}
	@Override
	int getCellIndex(CellAdv cell) {
		throw new UnsupportedOperationException("not implement");
	}
	@Override
	public void release() {
		throw new IllegalStateException("never link proxy object and call it's release");
	}
	@Override
	public void checkOrphan() {}
	
	@Override
	void onModelEvent(ModelEvent event) {}
}
