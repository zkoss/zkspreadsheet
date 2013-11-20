package org.zkoss.zss.ngmodel.impl;

import java.lang.ref.WeakReference;

import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

class CellProxyImpl extends AbstractCell{
	WeakReference<AbstractSheet> sheetRef;
	int rowIdx;
	int columnIdx;
	AbstractCell proxy;
	
	
	public CellProxyImpl(AbstractSheet sheet, int row,int column) {
		this.sheetRef = new WeakReference<AbstractSheet>(sheet);;
		this.rowIdx = row;
		this.columnIdx = column;
	}
	
	private AbstractSheet getSheet(){
		AbstractSheet sheet = sheetRef.get();
		if(sheet==null){
			throw new IllegalStateException("proxy target lost, you should't keep this instance");
		}
		return sheet;
	}
	
	private void loadProxy(){
		if(proxy==null){
			proxy = (CellImpl)getSheet().getCellAt(rowIdx, columnIdx, false);
			if(proxy!=null){
				sheetRef.clear();
			}
		}
	}

	public boolean isNull() {
		loadProxy();
		return proxy==null?true:proxy.isNull();
	}


	public CellType getType() {
		loadProxy();
		return proxy==null?CellType.BLANK:proxy.getType();
	}


	public int getRowIndex() {
		loadProxy();
		return proxy==null?rowIdx:proxy.getRowIndex();
	}


	public int getColumnIndex() {
		loadProxy();
		return proxy==null?columnIdx:proxy.getRowIndex();
	}


	public void setValue(Object value) {
		loadProxy();
		if(proxy==null){
			proxy = (CellImpl)((AbstractRow)getSheet().getOrCreateRowAt(rowIdx)).getOrCreateCellAt(columnIdx);
		}
		proxy.setValue(value);
	}


	public Object getValue() {
		loadProxy();
		return proxy==null?null:proxy.getValue();
	}

	public String asString(boolean enableSheetName) {
		loadProxy();
		return proxy==null?new CellReference(enableSheetName?getSheet().getSheetName():null, rowIdx,columnIdx,false,false).formatAsString():
			proxy.asString(enableSheetName);
	}

	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	public NCellStyle getCellStyle(boolean local) {
		loadProxy();
		if(proxy!=null){
			return proxy.getCellStyle(local);
		}
		if(local)
			return null;
		AbstractSheet sheet = getSheet();
		AbstractRow row = (AbstractRow)sheet.getRowAt(rowIdx, false);
		NCellStyle style = null;
		if(row!=null){
			style = row.getCellStyle(true);
		}
		if(style==null){
			ColumnImpl col = (ColumnImpl)sheet.getColumnAt(columnIdx, false);
			if(col!=null){
				style = col.getCellStyle(true);
			}
		}
		if(style==null){
			style = sheet.getBook().getDefaultCellStyle();
		}
		return style;
	}

	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		loadProxy();
		if(proxy==null){
			proxy = (CellImpl)((AbstractRow)getSheet().getOrCreateRowAt(rowIdx)).getOrCreateCellAt(columnIdx);
		}
		proxy.setCellStyle(cellStyle);
	}


}
