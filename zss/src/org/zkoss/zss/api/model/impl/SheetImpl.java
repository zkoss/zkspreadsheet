package org.zkoss.zss.api.model.impl;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Picture;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.DrawingManager;
import org.zkoss.zss.model.sys.impl.SheetCtrl;

public class SheetImpl implements Sheet{
	ModelRef<XSheet> sheetRef;
	Book nbook;
	public SheetImpl(ModelRef<XSheet> sheet){
		this.sheetRef = sheet;
	}
	
	public XSheet getNative(){
		return sheetRef.get();
	}
	public ModelRef<XSheet> getRef(){
		return sheetRef;
	}
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((sheetRef == null) ? 0 : sheetRef.hashCode());
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
		SheetImpl other = (SheetImpl) obj;
		if (sheetRef == null) {
			if (other.sheetRef != null)
				return false;
		} else if (!sheetRef.equals(other.sheetRef))
			return false;
		return true;
	}

	public Book getBook() {
		if(nbook!=null){
			return nbook;
		}
		nbook = new BookImpl(new SimpleRef<XBook>(getNative().getBook()));
		return nbook;
	}
	

	public boolean isProtected() {
		return getNative().getProtect();
	}

	public boolean isAutoFilterEnabled() {
		return getNative().isAutoFilterMode();
	}

	public boolean isDisplayGridlines() {
		return getNative().isDisplayGridlines();
	}

	public String getSheetName() {
		return getNative().getSheetName();
	}

	public boolean isRowHidden(int row) {
		final Row r = getNative().getRow(row);
		return r != null && r.getZeroHeight();
	}

	public boolean isColumnHidden(int column) {
		return getNative().isColumnHidden(column);
	}

	
	public List<Chart> getCharts(){
		Book book = getBook();
		DrawingManager dm = ((SheetCtrl)getNative()).getDrawingManager();
		List<Chart> charts = new ArrayList<Chart>();
		for(org.zkoss.poi.ss.usermodel.Chart chart:dm.getCharts()){
			charts.add(new ChartImpl(((BookImpl)book).getRef(), new SimpleRef<org.zkoss.poi.ss.usermodel.Chart>(chart)));
		}
		return charts;
	}

	
	public List<Picture> getPictures(){
		Book book = getBook();
		DrawingManager dm = ((SheetCtrl)getNative()).getDrawingManager();
		List<Picture> pictures = new ArrayList<Picture>();
		for(org.zkoss.poi.ss.usermodel.Picture pic:dm.getPictures()){
			pictures.add(new PictureImpl(((BookImpl)book).getRef(), new SimpleRef<org.zkoss.poi.ss.usermodel.Picture>(pic)));
		}
		return pictures;
	}

	public int getRowFreeze() {
		return BookHelper.getRowFreeze(getNative());
	}

	public int getColumnFreeze() {
		return BookHelper.getColumnFreeze(getNative());
	}

	@Override
	public boolean isPrintGridlines() {
		return getNative().isPrintGridlines();
	}

	@Override
	public org.zkoss.poi.ss.usermodel.Sheet getPoiSheet() {
		return getNative();
	}
	
}
