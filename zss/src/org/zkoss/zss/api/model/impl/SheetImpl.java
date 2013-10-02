/* SheetImpl.java

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

import java.util.ArrayList;
import java.util.List;

import org.zkoss.poi.hssf.usermodel.HSSFClientAnchor;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.xssf.usermodel.XSSFClientAnchor;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.UnitUtil;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Picture;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.model.sys.impl.DrawingManager;
import org.zkoss.zss.model.sys.impl.HSSFSheetImpl;
import org.zkoss.zss.model.sys.impl.SheetCtrl;
import org.zkoss.zss.ui.impl.XUtils;
/**
 * 
 * @author dennis
 * @since 3.0.0
 */
public class SheetImpl implements Sheet{
	private ModelRef<XSheet> _sheetRef;
	private ModelRef<XBook> _bookRef;
	private Book _book;
	public SheetImpl(ModelRef<XBook> book,ModelRef<XSheet> sheet){
		this._bookRef = book;
		this._sheetRef = sheet;
	}
	
	public XSheet getNative(){
		return _sheetRef.get();
	}
	public ModelRef<XSheet> getRef(){
		return _sheetRef;
	}
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((_sheetRef == null) ? 0 : _sheetRef.hashCode());
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
		if (_sheetRef == null) {
			if (other._sheetRef != null)
				return false;
		} else if (!_sheetRef.equals(other._sheetRef))
			return false;
		return true;
	}

	public Book getBook() {
		if(_book!=null){
			return _book;
		}
		_book = new BookImpl(_bookRef);
		return _book;
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
			charts.add(new ChartImpl(_sheetRef, new SimpleRef<org.zkoss.poi.ss.usermodel.Chart>(chart)));
		}
		return charts;
	}

	
	public List<Picture> getPictures(){
		DrawingManager dm = ((SheetCtrl)getNative()).getDrawingManager();
		List<Picture> pictures = new ArrayList<Picture>();
		for(org.zkoss.poi.ss.usermodel.Picture pic:dm.getPictures()){
			pictures.add(new PictureImpl(_sheetRef, new SimpleRef<org.zkoss.poi.ss.usermodel.Picture>(pic)));
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

	@Override
	public int getRowHeight(int row) {
		return XUtils.getRowHeightInPx(getNative(), row);
	}

	@Override
	public int getColumnWidth(int column) {
		return XUtils.getColumnWidthInPx(getNative(), column);
	}

	@Override
	public Object getSync() {
		return getNative();
	}
	
	
	/**
	 * Utility method, internal use only
	 */
	public static ClientAnchor toClientAnchor(org.zkoss.poi.ss.usermodel.Sheet sheet,SheetAnchor anchor){
		ClientAnchor can = null;
		if(sheet instanceof HSSFSheetImpl){//2003
			//it looks like not correct, need to double check this
			can = new HSSFClientAnchor(anchor.getXOffset(),anchor.getYOffset(),anchor.getLastXOffset(),anchor.getLastYOffset(),
					(short)anchor.getColumn(),(short)anchor.getRow(),(short)anchor.getLastColumn(),(short)anchor.getLastRow());
		}else{
			//code refer from ActionHandler.getClientAngent
			//but it looks like not correct, need to double check this
			can = new XSSFClientAnchor(UnitUtil.pxToEmu(anchor.getXOffset()),UnitUtil.pxToEmu(anchor.getYOffset()),
					UnitUtil.pxToEmu(anchor.getLastXOffset()),UnitUtil.pxToEmu(anchor.getLastYOffset()),
					(short)anchor.getColumn(),(short)anchor.getRow(),(short)anchor.getLastColumn(),(short)anchor.getLastRow());
//			can = new XSSFClientAnchor(anchor.getXOffset(),anchor.getYOffset(),anchor.getLastXOffset(),anchor.getLastYOffset(),
//					(short)anchor.getColumn(),(short)anchor.getRow(),(short)anchor.getLastColumn(),(short)anchor.getLastRow());
		}
		return can;
	}
	
	/**
	 * Utility method, internal use only
	 */
	public static SheetAnchor toSheetAnchor(org.zkoss.poi.ss.usermodel.Sheet sheet,ClientAnchor anchor){
		SheetAnchor san = null;
		if(sheet instanceof HSSFSheetImpl){
			san = new SheetAnchor(anchor.getRow1(),anchor.getCol1(),anchor.getDx1(),anchor.getDy1(),
					anchor.getRow2(),anchor.getCol2(),anchor.getDx2(),anchor.getDy2());
		}else{
			san = new SheetAnchor(anchor.getRow1(),anchor.getCol1(),UnitUtil.emuToPx(anchor.getDx1()),UnitUtil.emuToPx(anchor.getDy1()),
					anchor.getRow2(),anchor.getCol2(),UnitUtil.emuToPx(anchor.getDx2()),UnitUtil.emuToPx(anchor.getDy2()));
		}
		return san;
	}

	@Override
	public int getFirstRow() {
		return getPoiSheet().getFirstRowNum();
	}

	@Override
	public int getLastRow() {
		return getPoiSheet().getLastRowNum();
	}

	@Override
	public int getFirstColumn(int row) {
		Row r = getPoiSheet().getRow(row);
		return r==null?-1:r.getFirstCellNum();
	}

	@Override
	public int getLastColumn(int row) {
		Row r = getPoiSheet().getRow(row);
		return r==null?-1:r.getLastCellNum();
	}
}
