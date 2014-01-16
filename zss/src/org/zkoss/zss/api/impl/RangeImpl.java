/* RangeImpl.java

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
package org.zkoss.zss.api.impl;

import java.util.concurrent.locks.ReadWriteLock;

import org.zkoss.poi.ss.formula.FormulaParseException;
import org.zkoss.zss.api.*;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.Chart.Type;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.Picture.Format;
import org.zkoss.zss.api.model.impl.*;
import org.zkoss.zss.ngapi.*;
import org.zkoss.zss.ngmodel.*;

/**
 * 1.Range is not handling the protection issue, if you have handle it yourself before calling the api(by calling {@code #isProtected()})
 * @author dennis
 * @since 3.0.0
 */
public class RangeImpl implements Range{
	
	private NRange _range;
	
	private CellStyleHelper _cellStyleHelper;
	private CellData _cellData;
	
	/**
	 * @deprecated since 3.5 it is always synchronized on book by a read write lock.
	 */
	public void setSyncLevel(SyncLevel syncLevel){
	}
	
	private SharedContext _sharedCtx;
	
	public RangeImpl(NRange range,Sheet sheet) {
		this._range = range;
		_sharedCtx = new SharedContext(sheet);
	}
	private RangeImpl(NRange range,SharedContext ctx) {
		this._range = range;
		_sharedCtx = ctx;
	}
	
	public ReadWriteLock getLock(){
		return _range.getLock();
	}
	
	public CellStyleHelper getCellStyleHelper(){
		if(_cellStyleHelper==null){
			_cellStyleHelper = new CellStyleHelperImpl(getBook());
		}
		return _cellStyleHelper;
	}
	
	public CellData getCellData(){
		if(_cellData==null){
			_cellData = new CellDataImpl(this);
		}
		return _cellData;
	}
	
	public NRange getNative(){
		return _range;
	}

	
	private static class SharedContext{
		Sheet _sheet;
		
		private SharedContext(Sheet sheet){
			this._sheet = sheet;
		}
		
		public Sheet getSheet(){
			return _sheet;
		}
		
		public Book getBook(){
			return _sheet.getBook();
		}
	}
	
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((_range == null) ? 0 : _range.hashCode());
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
		RangeImpl other = (RangeImpl) obj;
		if (_range == null) {
			if (other._range != null)
				return false;
		} else if (!_range.equals(other._range))
			return false;
		return true;
	}

	public boolean isProtected() {
		return _range.isSheetProtected();
	}	

	public boolean isAnyCellProtected(){
		return _range.isAnyCellProtected();
	}
	
	public Range paste(Range dest, boolean cut) { 
		NRange r = _range.copy(((RangeImpl)dest).getNative(), cut);
		return new RangeImpl(r, dest.getSheet());
	}

	/* short-cut for pasteSpecial, it is original Range.copy*/
	public Range paste(Range dest) { 
		NRange r = _range.copy(((RangeImpl)dest).getNative());
		return new RangeImpl(r, dest.getSheet());
	}
	
	public Range pasteSpecial(Range dest,PasteType type,PasteOperation op,boolean skipBlanks,boolean transpose) { 
		NRange r = _range.pasteSpecial(((RangeImpl)dest).getNative(), EnumUtil.toRangePasteTypeNative(type), EnumUtil.toRangePasteOpNative(op), skipBlanks, transpose);
		return new RangeImpl(r, dest.getSheet());
	}


	public void clearContents() { 
		_range.clearContents();
	}
	
	public Sheet getSheet(){
		return _sharedCtx.getSheet();
	}

 
	public void clearStyles() {
		setCellStyle(null);//will use default book cell style		
	}

	public void setCellStyle(final CellStyle nstyle) { 
		_range.setStyle(nstyle==null?null:((CellStyleImpl)nstyle).getNative());
	}


	public int getColumn() {
		return _range.getColumn();
	}
	public int getRow() {
		return _range.getRow();
	}
	public int getLastColumn() {
		return _range.getLastColumn();
	}
	public int getLastRow() {
		return _range.getLastRow();
	}
	
	public void sync(RangeRunner run){
		ReadWriteLock lock = _range.getLock();
		lock.writeLock().lock();
		try{
			run.run(this);
		}finally{
			lock.writeLock().unlock();
		}
		return;
	}
	/**
	 * visit all cells in this range, make sure you call this in a limited range, 
	 * don't use it for all row/column selection, it will spend much time to iterate the cell 
	 * @param visitor the visitor 
	 * @param create create cell if it doesn't exist, if it is true, it will also lock the sheet
	 */
	public void visit(final CellVisitor visitor){
		final int r=getRow();
		final int lr=getLastRow();
		final int c=getColumn();
		final int lc=getLastColumn();
		
		Runnable run = new Runnable(){
			public void run(){
				for(int i=r;i<=lr;i++){
					for(int j=c;j<=lc;j++){
						if(!visitCell(visitor,i,j))
							break;
					}
				}
			}
		};
		
		ReadWriteLock lock = _range.getLock();
		lock.writeLock().lock();
		try{
			run.run();
		}finally{
			lock.writeLock().unlock();
		}
	}
	
	private boolean visitCell(CellVisitor visitor,int r, int c){
		boolean ignore = false;
		NSheet sheet = _range.getSheet();
		NCell cell = sheet.getCell(r,c);
		if(cell.isNull()){
			if(ignore){
				return true;
			}else{
				//use less call, to just compatible with 3.0
				visitor.createIfNotExist(r, c);
			}
		}
		return visitor.visit(new RangeImpl(NRanges.range(_range.getSheet(),r,c),_sharedCtx));
	}

	public Book getBook() {
		return getSheet().getBook();
	}
	
	public void applyBordersAround(BorderType borderType,String htmlColor){
		applyBorders(ApplyBorderType.OUTLINE,borderType, htmlColor);
	}
	
	public void applyBorders(ApplyBorderType type,BorderType borderType,String htmlColor){ 
		_range.setBorders(EnumUtil.toRangeApplyBorderType(type), EnumUtil.toRangeBorderType(borderType), htmlColor);
	}

	
	public boolean hasMergedCell(){
		CellRegion curr = new CellRegion(getRow(),getColumn(),getLastRow(),getLastColumn());
		for(CellRegion ma:_range.getSheet().getMergedRegions()){
			if(curr.overlaps(ma)){
				return true;
			}
		}
		return false;
	}
	
	public boolean isMergedCell(){
		for(CellRegion ma:_range.getSheet().getMergedRegions()){
			if(ma.equals(getRow(),getColumn(),getLastRow(),getLastColumn())){
				return true;
			}
		}
		return false;
	}
	
	
	static private class Result<T> {
		T r;
		Result(){}
		Result(T r){
			this.r = r;
		}
		
		public T get(){
			return r;
		}
		
		public void set(T r){
			this.r = r;
		}
	}

	public void merge(boolean across){ 
		_range.merge(across);
	}
	
	public void unmerge(){ 
		_range.unmerge();
	}

	
	public RangeImpl toShiftedRange(int rowOffset,int colOffset){ 
		RangeImpl offsetRange = new RangeImpl(_range.getOffset(rowOffset, colOffset),_sharedCtx);
		return offsetRange;
	}
	
	
	public RangeImpl toCellRange(int rowOffset,int colOffset){
		RangeImpl cellRange = new RangeImpl(NRanges.range(_range.getSheet(),getRow()+rowOffset,getColumn()+colOffset),_sharedCtx);
		return cellRange;
	}
	
	/** get the top-left cell range of this range**/
	public RangeImpl getLeftTop() { 
		return toCellRange(0,0);
	}
	
	/**
	 *  Return a range that represents all columns and between the first-row and last-row of this range
	 **/
	public RangeImpl toRowRange(){ 
		return new RangeImpl(_range.getRows(),_sharedCtx);
	}
	
	/**
	 *  Return a range that represents all rows and between the first-column and last-column of this range
	 **/
	public RangeImpl toColumnRange(){ 
		return new RangeImpl(_range.getColumns(),_sharedCtx);
	}
	
	/**
	 * Check if this range represents a whole column, which mean all rows are included, 
	 */
	public boolean isWholeColumn(){ 
		return _range.isWholeColumn();
	}
	/**
	 * Check if this range represents a whole row, which mean all column are included, 
	 */
	public boolean isWholeRow(){ 
		return _range.isWholeRow();
	}
	/**
	 * Check if this range represents a whole sheet, which mean all column and row are included, 
	 */
	public boolean isWholeSheet(){ 
		return _range.isWholeSheet();
	}
	
	public void insert(InsertShift shift,InsertCopyOrigin copyOrigin){
		_range.insert(EnumUtil.toRangeInsertShift(shift), EnumUtil.toRangeInsertCopyOrigin(copyOrigin));
	}
	
	public void delete(DeleteShift shift){ 
		_range.delete(EnumUtil.toRangeDeleteShift(shift));
	}
	
	public void sort(boolean desc){	
		sort(desc,false,false,false,null);
	}
	
	public void sort(boolean desc,
			boolean header, 
			boolean matchCase, 
			boolean sortByRows, 
			SortDataOption dataOption){
		Range index = null;
		int r = getRow();
		int c = getColumn();
		int lr = getLastRow();
		int lc = getLastColumn();
		
		index = Ranges.range(this.getSheet(),r,c,sortByRows?r:lr,sortByRows?lc:c);
		
		sort(index,desc,dataOption,
			null,false,null,
			null,false,null,
			header,matchCase,sortByRows);
	}
	
	public void sort(Range index1,boolean desc1,SortDataOption dataOption1,
			Range index2,boolean desc2,SortDataOption dataOption2,
			Range index3,boolean desc3,SortDataOption dataOption3,
			boolean header, 
			/*int orderCustom, //not implement*/
			boolean matchCase, 
			boolean sortByRows
			/*int sortMethod, //not implement*/){
		
		//TODO review the full impl for range1,range2,range3
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 */
//		_range.sort(index1==null?null:((RangeImpl)index1).getNative(), desc1, 
//				index2==null?null:((RangeImpl)index2).getNative()/*rng2*/, -1 /*type*/, desc2/*desc2*/, 
//				index3==null?null:((RangeImpl)index3).getNative()/*rng3*/, desc3/*desc3*/,
//				header?BookHelper.SORT_HEADER_YES:BookHelper.SORT_HEADER_NO/*header*/,
//				-1/*orderCustom*/, matchCase, sortByRows, -1/*sortMethod*/, 
//				dataOption1==null?BookHelper.SORT_NORMAL_DEFAULT:EnumUtil.toRangeSortDataOption(dataOption1)/*dataOption1*/,
//				dataOption2==null?BookHelper.SORT_NORMAL_DEFAULT:EnumUtil.toRangeSortDataOption(dataOption2)/*dataOption2*/,
//				dataOption3==null?BookHelper.SORT_NORMAL_DEFAULT:EnumUtil.toRangeSortDataOption(dataOption3)/*dataOption3*/);
		
	}
	
	/** check if auto filter is enable or not.**/
	public boolean isAutoFilterEnabled(){
		return getSheet().isAutoFilterEnabled();
	}
	
	// ZSS-246: give an API for user checking the auto-filtering range before applying it.
	public Range findAutoFilterRange() {
		NRange r = _range.findAutoFilterRange();
		if(r != null) {
			return Ranges.range(getSheet(), r.getRow(), r.getColumn(), r.getLastRow(), r.getLastColumn());
		} else {
			return null;
		}
	}

	/** enable/disable autofilter of the sheet**/
	public void enableAutoFilter(boolean enable){
		if(isAutoFilterEnabled() == enable){
			return ;
		} 
		_range.enableAutoFilter(enable);
	}
	/** enable filter with condition **/
	//TODO have to review this after I know more detail
	public void enableAutoFilter(int field, AutoFilterOperation filterOp, Object criteria1, Object criteria2, Boolean showButton){ 
		_range.enableAutoFilter(field,EnumUtil.toRangeAutoFilterOperation(filterOp),criteria1,criteria2,showButton);
	}
	
	/** clear criteria of all filters, show all the data**/
	public void resetAutoFilter(){
		_range.resetAutoFilter();
	}
	
	/** re-apply existing criteria of filters **/
	public void applyAutoFilter(){ 
		_range.applyAutoFilter();
	}
	
	/** enable sheet protection and apply a password**/
	public void protectSheet(String password){ 
		_range.protectSheet(password);
	}
	
	public void autoFill(Range dest,AutoFillType fillType){
		//TODO the syncLevel
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 
		_range.autoFill(((RangeImpl)dest).getNative(), EnumUtil.toRangeAutoFillType(fillType));
		*/
	}
	
	public void fillDown(){ 
		_range.fillDown();
	}
	
	public void fillLeft(){ 
		_range.fillLeft();
	}
	
	public void fillUp(){ 
		_range.fillUp();
	}
	
	public void fillRight(){ 
		_range.fillRight();
	}
	
	/** shift this range with a offset row and column**/
	public void shift(int rowOffset,int colOffset){ 
		_range.move(rowOffset, colOffset);
	}
	
	public String getCellEditText(){
		String txt = _range.getEditText();
		return txt==null?"":txt;
	}
	
	public void setCellEditText(String editText){ 
		try{
			_range.setEditText(editText);
			//TODO zss 3.5 the parse exception
		}catch(FormulaParseException x){
			throw new IllegalFormulaException(x.getMessage(),x);
		}
	}
	
	public String getCellFormatText(){ 
		return _range.getCellFormatText();
	}
	
	public Object getCellValue(){ 
		return _range.getValue();
	}
	
	public void setDisplaySheetGridlines(boolean enable){ 
		_range.setDisplayGridlines(enable);
	}
	
	public boolean isDisplaySheetGridlines(){ 
		return getSheet().isDisplayGridlines();
	}
	
	public void setHidden(boolean hidden){ 
		_range.setHidden(hidden);
	}
	
	public void setCellHyperlink(HyperlinkType type,String address,String display){ 
		_range.setHyperlink(EnumUtil.toHyperlinkType(type), address, display);
	}
	
	public Hyperlink getCellHyperlink(){ 
		NHyperlink l = _range.getHyperlink();
		return l==null?null:new HyperlinkImpl(new SimpleRef<NHyperlink>(l),getCellEditText());
	}
	
	public void setSheetName(String name){ 
		_range.setSheetName(name);
	}
	
	public String getSheetName(){
		return getSheet().getSheetName();
	}
	
	public void setSheetOrder(int pos){ 
		_range.setSheetOrder(pos);
	}
	
	public int getSheetOrder(){
		return getBook().getSheetIndex(getSheet());
	}
	
	public void setCellValue(Object value){
		_range.setValue(value);
	}
	
	private ModelRef<NBook> getBookRef(){
		return ((BookImpl)getBook()).getRef();
	}
	
	private ModelRef<NSheet> getSheetRef(){
		return ((SheetImpl)getSheet()).getRef();
	}
	

	/**
	 * get the first cell style of this range
	 * 
	 * @return cell style if cell is exist, the check row style and column cell style if cell not found, if row and column style is not exist, then return default style of sheet
	 */
	public CellStyle getCellStyle() {
		NSheet sheet = _range.getSheet();
		NBook book = sheet.getBook();
		
		int r = _range.getRow();
		int c = _range.getColumn();
		NCellStyle style = sheet.getCell(r, c).getCellStyle();

		return new CellStyleImpl(getBookRef(), new SimpleRef<NCellStyle>(style));		
	}

	
	public Picture addPicture(SheetAnchor anchor,byte[] image,Format format){
		NPicture picture = _range.addPicture(SheetImpl.toViewAnchor(_range.getSheet(), anchor), image, EnumUtil.toPictureFormat(format));
		return new PictureImpl(new SimpleRef<NSheet>(_range.getSheet()), new SimpleRef<NPicture>(picture));
	}
	
	public void deletePicture(Picture picture){
		_range.deletePicture(((PictureImpl)picture).getNative());
	}
	
	public void movePicture(SheetAnchor anchor,Picture picture){
		_range.movePicture(((PictureImpl)picture).getNative(), SheetImpl.toViewAnchor(_range.getSheet(), anchor));
	}
	
	//currently, we only support to modify chart in XSSF
	public Chart addChart(SheetAnchor anchor,ChartData data,Type type, Grouping grouping, LegendPosition pos){
		//TODO the syncLevel
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 
		ClientAnchor an = SheetImpl.toClientAnchor(getSheet().getPoiSheet(),anchor);
		org.zkoss.poi.ss.usermodel.charts.ChartData cdata = ((ChartDataImpl)data).getNative();
		org.zkoss.poi.ss.usermodel.Chart chart = _range.addChart(an, cdata, EnumUtil.toChartType(type), EnumUtil.toChartGrouping(grouping), EnumUtil.toLegendPosition(pos));
		return new ChartImpl(getSheetRef(), new SimpleRef<org.zkoss.poi.ss.usermodel.Chart>(chart));
		*/
	}
	
	//currently, we only support to modify chart in XSSF
	public void deleteChart(Chart chart){
		//TODO the syncLevel
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 
		_range.deleteChart(((ChartImpl)chart).getNative());
		*/
	}
	
	//currently, we only support to modify chart in XSSF
	public void moveChart(SheetAnchor anchor,Chart chart){
		//TODO the syncLevel
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 
		ClientAnchor an = SheetImpl.toClientAnchor(getSheet().getPoiSheet(),anchor);
		_range.moveChart(((ChartImpl)chart).getNative(), an);
		*/
	}
	
	
	public Sheet createSheet(String name){
		NSheet sheet = _range.createSheet(name);
		return new SheetImpl(getBookRef(),new SimpleRef(sheet));
	}
	
	public void deleteSheet(){
		_range.deleteSheet();
	}
	
	
	@Override
	public void setColumnWidth(int widthPx) {
		NRange r = _range.isWholeColumn()?_range:_range.getColumns();
		r.setColumnWidth(widthPx);
	}
	@Override
	public void setRowHeight(int heightPx) {
		NRange r = _range.isWholeRow()?_range:_range.getRows();
		r.setRowHeight(heightPx);
	}
	
	//api that need special object wrap
	
	
	private void apiSpecialWrapObject(){

//		range.getFormatText();//FormatText
//		range.getHyperlink();//Hyperlink
//		
//		range.getRichEditText();//RichTextString
//		range.getText();//RichTextString (what is the difference of getRichEditText)
//		
		
//		range.validate("");//DataValidation
		
	}
	
	
	private void api4Internal(){
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 
		_range.notifyDeleteFriendFocus(null);//by Spreadsheet
		_range.notifyMoveFriendFocus(null);//
		*/
	}
	
	
	//API of range that no one use it.
	
	
	private void apiNoOneUse(){
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 
		_range.getCount();
		_range.getCurrentRegion();
		_range.getDependents();
		_range.getDirectDependents();
		_range.getPrecedents();
		
		_range.isCustomHeight();
		
		//range.pasteSpecial(pasteType, pasteOp, SkipBlanks, transpose);
		 */
	}
	
	public String toString(){
		return Ranges.getAreaRefString(getSheet(), getRow(),getColumn(),getLastRow(),getLastColumn());
	}
	
	/**
	 * Notify this range has been changed.
	 */
	public void notifyChange(){ 
		_range.notifyChange();
	}
	
	public void notifyChange(String[] variables){
		throw new UnsupportedOperationException("not implment yet");/* zss 3.5 
		((XBook)getBook().getPoiBook()).notifyChange(variables);
		*/
	}
	
	@Override
	public void setFreezePanel(int rowfreeze, int columnfreeze) { 
		_range.setFreezePanel(rowfreeze, columnfreeze);
	}
	@Override
	public int getRowCount() {
		return _range.getLastRow()-_range.getRow()+1;
	}
	@Override
	public int getColumnCount() {
		return _range.getLastColumn()-_range.getColumn()+1;
	}
	@Override
	public String asString() {
		return Ranges.getAreaRefString(getSheet(), getRow(),getColumn(),getLastRow(),getLastColumn());
	}
		

}
