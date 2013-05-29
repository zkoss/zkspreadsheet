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

import java.util.ArrayList;
import java.util.List;

import org.zkoss.poi.hssf.usermodel.HSSFClientAnchor;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFClientAnchor;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.CellVisitor;
import org.zkoss.zss.api.RangeRunner;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.SheetAnchor;
import org.zkoss.zss.api.UnitUtil;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Book.BookType;
import org.zkoss.zss.api.model.CellData;
import org.zkoss.zss.api.model.CellStyle;
import org.zkoss.zss.api.model.CellStyle.BorderType;
import org.zkoss.zss.api.model.Chart;
import org.zkoss.zss.api.model.Chart.Grouping;
import org.zkoss.zss.api.model.Chart.LegendPosition;
import org.zkoss.zss.api.model.Chart.Type;
import org.zkoss.zss.api.model.ChartData;
import org.zkoss.zss.api.model.Hyperlink;
import org.zkoss.zss.api.model.Hyperlink.HyperlinkType;
import org.zkoss.zss.api.model.Picture;
import org.zkoss.zss.api.model.Picture.Format;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.impl.CellDataImpl;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.api.model.impl.HyperlinkImpl;
import org.zkoss.zss.api.model.impl.ModelRef;
import org.zkoss.zss.api.model.impl.BookImpl;
import org.zkoss.zss.api.model.impl.CellStyleImpl;
import org.zkoss.zss.api.model.impl.ChartDataImpl;
import org.zkoss.zss.api.model.impl.ChartImpl;
import org.zkoss.zss.api.model.impl.PictureImpl;
import org.zkoss.zss.api.model.impl.SheetImpl;
import org.zkoss.zss.api.model.impl.SimpleRef;
import org.zkoss.zss.model.sys.XBook;
import org.zkoss.zss.model.sys.XRange;
import org.zkoss.zss.model.sys.XRanges;
import org.zkoss.zss.model.sys.XSheet;
import org.zkoss.zss.model.sys.impl.BookHelper;
import org.zkoss.zss.ui.impl.XUtils;

/**
 * 1.Range is not handling the protection issue, if you have handle it yourself before calling the api(by calling {@code #isProtected()})
 * @author dennis
 * @since 3.0.0
 */
public class RangeImpl implements Range{
	
	XRange range;
	
	SyncLevel syncLevel = SyncLevel.BOOK;
	
	CellStyleHelper cellStyleHelper;
	CellData cellData;
	
	public void setSyncLevel(SyncLevel syncLevel){
		this.syncLevel = syncLevel;
	}
	
	private SharedContext sharedCtx;
	
	public RangeImpl(Sheet sheet,XRange range) {
		this.range = range;
		sharedCtx = new SharedContext((SheetImpl)sheet);
	}
	private RangeImpl(XRange range,SharedContext ctx) {
		this.range = range;
		sharedCtx = ctx;
	}
	
	
	public CellStyleHelper getCellStyleHelper(){
		if(cellStyleHelper==null){
			cellStyleHelper = new CellStyleHelperImpl(this);
		}
		return cellStyleHelper;
	}
	
	public CellData getCellData(){
		if(cellData==null){
			cellData = new CellDataImpl(this);
		}
		return cellData;
	}
	
	public XRange getNative(){
		return range;
	}
	
	static class SharedContext{
		SheetImpl nsheet;
		List<MergeArea> mergeAreas;
		
		private SharedContext(SheetImpl nsheet){
			this.nsheet = nsheet;
			
		}
		
		public Sheet getSheet(){
			return nsheet;
		}
		
		
		public List<MergeArea> getMergeAreas(){
			initMergeRangesCache();
			return mergeAreas;
		}
		
		
		private void initMergeRangesCache(){
			if(mergeAreas==null){//TODO a better way to cache(index) this.(I this MergeMatrixHelper is not good enough now)
				XSheet sheet = nsheet.getNative();
				int sz = sheet.getNumMergedRegions();
				mergeAreas = new ArrayList<MergeArea>(sz);
				for(int j = sz - 1; j >= 0; --j) {
					final CellRangeAddress addr = sheet.getMergedRegion(j);
					mergeAreas.add(new MergeArea(addr.getFirstRow(),addr.getFirstColumn(), addr.getLastRow(),addr.getLastColumn()));
				}
			}
		}
		
		public void resetMergeAreas(){
			mergeAreas = null;
		}
	}
	private static class MergeArea{
		int r,lr,c,lc;
		public MergeArea(int r,int c,int lr,int lc){
			this.c = c;
			this.r = r;
			this.lc = lc;
			this.lr = lr;
		}

		public boolean contains(int row,int col) {
			return col >= this.c && col <= this.lc && row >= this.r && row <= this.lr;
		}

		public boolean isUnion(int row, int column, int lastRow, int lastColumn) {
			return contains(row, column) || contains(lastRow, lastColumn)
					|| contains(row, lastColumn) || contains(lastRow, column);
		}
	}
	
	
	@Override
	public int hashCode() {
		final int prime = 31;
		int result = 1;
		result = prime * result + ((range == null) ? 0 : range.hashCode());
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
		if (range == null) {
			if (other.range != null)
				return false;
		} else if (!range.equals(other.range))
			return false;
		return true;
	}

	public boolean isProtected() {
		return getSheet().isProtected();
	}	

//	public boolean isAnyCellProtected(){
//		return range.isAnyCellProtected();
//	}

	/* short-cut for pasteSpecial, it is original Range.copy*/
	/**
	 * @param dest the destination 
	 * @return true if paste successfully, past to a protected sheet with any
	 *         locked cell in the destination range will always cause past fail.
	 */
	public boolean paste(Range dest) {		
		return pasteSpecial(dest,PasteType.ALL,PasteOperation.NONE,false,false);
	}
	
	/**
	 * @param dest the destination 
	 * @param transpose TODO
	 * @return true if paste successfully, past to a protected sheet with any
	 *         locked cell in the destination range will always cause past fail.
	 */
	public boolean pasteSpecial(Range dest,PasteType type,PasteOperation op,boolean skipBlanks,boolean transpose) {
//		if(!isAnyCellProtected()){ // ranges seems this in copy/paste already
		//TODO the syncLevel
		XRange r = range.pasteSpecial(((RangeImpl)dest).getNative(), EnumUtil.toRangePasteTypeNative(type), EnumUtil.toRangePasteOpNative(op), skipBlanks, transpose);
		return r!=null;
//		}
	}


	public void clearContents() {
		//TODO the syncLevel
		range.clearContents();		
	}
	
	public Sheet getSheet(){
		return sharedCtx.getSheet();
	}

 
	public void clearStyles() {
		setCellStyle(null);//will use default book cell style		
	}

	public void setCellStyle(final CellStyle nstyle) {
		//TODO the syncLevel
		range.setStyle(nstyle==null?null:((CellStyleImpl)nstyle).getNative());
	}


	public int getColumn() {
		return range.getColumn();
	}
	public int getRow() {
		return range.getRow();
	}
	public int getLastColumn() {
		return range.getLastColumn();
	}
	public int getLastRow() {
		return range.getLastRow();
	}
	
	public void sync(RangeRunner run){
		switch(syncLevel){
		case NONE:
			run.run(this);
			return;
		case BOOK:
			synchronized(range.getSheet().getBook()){//it just show concept, we have a betterway to do read-write lock
				run.run(this);
			}
			return;
		}
	}
	/**
	 * visit all cells in this range, make sure you call this in a limited range, 
	 * don't use it for all row/column selection, it will spend much time to iterate the cell 
	 * @param visitor the visitor 
	 * @param create create cell if it doesn't exist, if it is true, it will also lock the sheet
	 */
	public void visit(CellVisitor visitor){
		visit0(visitor,syncLevel);
	}
	private void visit0(final CellVisitor visitor,SyncLevel sync){
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
		
		switch(sync){
		case NONE:
			run.run();
			return;
		case BOOK:
			synchronized(range.getSheet().getBook()){
				run.run();
			}
			return;
		}
	}
	
	private boolean visitCell(CellVisitor visitor,int r, int c){
		boolean ignore = false;
		boolean ignoreSet = false;
		boolean create = false;
		boolean createSet = false;
		XSheet sheet = range.getSheet();
		Row row = sheet.getRow(r);
		if(row==null){
			ignore = visitor.ignoreIfNotExist(r,c);
			ignoreSet = true;
			if(!ignore){
				create = visitor.createIfNotExist(r, c);
				createSet = true;
				if(create){
					row = sheet.createRow(r);
				}
			}else{
				return true;
			}
		}
		Cell cell = row.getCell(c);
		if(cell==null){
			if(!ignoreSet){
				ignore = visitor.ignoreIfNotExist(r,c);
				ignoreSet = true;
			}
			if(!ignore){
				if(!createSet){
					create = visitor.createIfNotExist(r, c);
					createSet = true;
				}
				if(create){
					cell = row.createCell(c);
				}
			}else{
				return true;
			}
		}
		return visitor.visit(new RangeImpl(XRanges.range(range.getSheet(),r,c),sharedCtx));
	}

	public Book getBook() {
		return getSheet().getBook();
	}
	
	public void applyBordersAround(BorderType borderType,String htmlColor){
		applyBorders(ApplyBorderType.OUTLINE,borderType, htmlColor);
	}
	
	public void applyBorders(ApplyBorderType type,BorderType borderType,String htmlColor){
		//TODO the syncLevel
		range.setBorders(EnumUtil.toRangeApplyBorderType(type), EnumUtil.toRangeBorderType(borderType), htmlColor);
	}

	
	public boolean hasMergedCell(){
		final Result<Boolean> result = new Result<Boolean>(Boolean.FALSE);
		//TODO use visitor is bad performance, I should check merge area directly
		for(MergeArea ma:sharedCtx.getMergeAreas()){
			if(ma.isUnion(getRow(),getColumn(),getLastRow(),getLastColumn())){
				return true;
			}
		}
		return false;
		
		
//		visit0(new CellVisitor(){
//			public boolean ignoreIfNotExist(int row, int column) {
//				return true;
//			}
//			@Override
//			public boolean visit(Range cellRange) {
//				int c = cellRange.getColumn();
//				int r = cellRange.getRow();
//				for(MergeArea ma:sharedCtx.getMergeAreas()){//TODO index, not loop
//					if(ma.contains(c, r)){
//						result.set(Boolean.TRUE);
//						return false;//break visit
//					}
//				}
//				return true;
//			}},SyncLevel.NONE);
//		return result.get();
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
		//TODO the syncLevel
		range.merge(across);
	}
	
	public void unMerge(){
		//TODO the syncLevel
		range.unMerge();
	}

	
	public RangeImpl toShiftedRange(int rowOffset,int colOffset){
		RangeImpl offsetRange = new RangeImpl(range.getOffset(rowOffset, colOffset),sharedCtx);
		return offsetRange;
	}
	
	
	public RangeImpl toCellRange(int rowOffset,int colOffset){
		RangeImpl cellRange = new RangeImpl(XRanges.range(range.getSheet(),getRow()+rowOffset,getColumn()+colOffset),sharedCtx);
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
		return new RangeImpl(range.getRows(),sharedCtx);
	}
	
	/**
	 *  Return a range that represents all rows and between the first-column and last-column of this range
	 **/
	public RangeImpl toColumnRange(){
		return new RangeImpl(range.getColumns(),sharedCtx);
	}
	
	/**
	 * Check if this range represents a whole column, which mean all rows are included, 
	 */
	public boolean isWholeColumn(){
		return range.isWholeColumn();
	}
	/**
	 * Check if this range represents a whole row, which mean all column are included, 
	 */
	public boolean isWholeRow(){
		return range.isWholeRow();
	}
	/**
	 * Check if this range represents a whole sheet, which mean all column and row are included, 
	 */
	public boolean isWholeSheet(){
		return range.isWholeSheet();
	}
	
	public void insert(InsertShift shift,InsertCopyOrigin copyOrigin){
		//TODO the syncLevel
		range.insert(EnumUtil.toRangeInsertShift(shift), EnumUtil.toRangeInsertCopyOrigin(copyOrigin));
	}
	
	public void delete(DeleteShift shift){
		//TODO the syncLevel
		range.delete(EnumUtil.toRangeDeleteShift(shift));
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
		
		sort(index,desc,header,matchCase,sortByRows,dataOption,
				null,false,null,null,false,null);
	}
	
	public void sort(Range index1,
			boolean desc1,
			boolean header, 
			/*int orderCustom, //not implement*/
			boolean matchCase, 
			boolean sortByRows, 
			/*int sortMethod, //not implement*/
			SortDataOption dataOption1,
			Range index2,boolean desc2,SortDataOption dataOption2,
			Range index3,boolean desc3,SortDataOption dataOption3){
		
		//TODO the syncLevel
		
		//TODO review the full impl for range1,range2,range3
		
		range.sort(index1==null?null:((RangeImpl)index1).getNative(), desc1, 
				index2==null?null:((RangeImpl)index1).getNative()/*rng2*/, -1 /*type*/, desc2/*desc2*/, 
				index3==null?null:((RangeImpl)index1).getNative()/*rng3*/, desc3/*desc3*/,
				header?BookHelper.SORT_HEADER_YES:BookHelper.SORT_HEADER_NO/*header*/,
				-1/*orderCustom*/, matchCase, sortByRows, -1/*sortMethod*/, 
				EnumUtil.toRangeSortDataOption(dataOption1)/*dataOption1*/,
				EnumUtil.toRangeSortDataOption(dataOption2)/*dataOption2*/,
				EnumUtil.toRangeSortDataOption(dataOption3)/*dataOption3*/);
	}
	
	/** check if auto filter is enable or not.**/
	public boolean isAutoFilterEnabled(){
		return getSheet().isAutoFilterEnabled();
	}
	
	/** enable/disable autofilter of the sheet**/
	public void enableAutoFilter(boolean enable){
		//TODO the syncLevel
		if(isAutoFilterEnabled() == enable){
			return ;
		}
		
		range.autoFilter();//toggle on/off automatically
	}
	/** enable filter with condition **/
	//TODO have to review this after I know more detail
	public void enableAutoFilter(int field, AutoFilterOperation filterOp, Object criteria1, Object criteria2, Boolean visibleDropDown){
		//TODO the syncLevel
		range.autoFilter(field,criteria1,EnumUtil.toRangeAutoFilterOperation(filterOp),criteria2,visibleDropDown);
	}
	
	/** clear condition of filter, show all the data**/
	public void resetAutoFilter(){
		//TODO the syncLevel
		range.showAllData();
	}
	/** apply the filter to filter again**/
	public void applyAutoFilter(){
		//TODO the syncLevel
		range.applyFilter();
	}
	
	/** enable sheet protection and apply a password**/
	public void protectSheet(String password){
		//TODO the syncLevel
		range.protectSheet(password);
	}
	
	public void autoFill(Range dest,AutoFillType fillType){
		//TODO the syncLevel
		range.autoFill(((RangeImpl)dest).getNative(), EnumUtil.toRangeAutoFillType(fillType));
	}
	
	public void fillDown(){
		//TODO the syncLevel
		range.fillDown();
	}
	
	public void fillLeft(){
		//TODO the syncLevel
		range.fillLeft();
	}
	
	public void fillUp(){
		//TODO the syncLevel
		range.fillUp();
	}
	
	public void fillRight(){
		//TODO the syncLevel
		range.fillRight();
	}
	
	/** shift this range with a offset row and column**/
	public void shift(int rowOffset,int colOffset){
		//TODO the syncLevel
		range.move(rowOffset, colOffset);
	}
	
	public String getCellEditText(){
		return range.getEditText();
	}
	
	public void setCellEditText(String editText){
		//TODO the syncLevel
		range.setEditText(editText);
	}
	
	public String getCellFormatText(){
		//I don't create my way, use the same way from Spreadsheet implementation as possible
		return XUtils.getCellFormatText(getNative().getSheet(), getRow(), getColumn());
	}
	
	//TODO need to verify the object type
	public Object getCellValue(){
		return range.getValue();
	}
	
	public void setDisplaySheetGridlines(boolean enable){
		//TODO the syncLevel
		range.setDisplayGridlines(enable);
	}
	
	public boolean isDisplaySheetGridlines(){
		return getSheet().isDisplayGridlines();
	}
	
	public void setHidden(boolean hidden){
		//TODO the syncLevel
		range.setHidden(hidden);
	}
	
	public void setCellHyperlink(HyperlinkType type,String address,String display){
		//TODO the syncLevel
		range.setHyperlink(EnumUtil.toHyperlinkType(type), address, display);
	}
	
	public Hyperlink getCellHyperlink(){
		org.zkoss.poi.ss.usermodel.Hyperlink l = range.getHyperlink();
		return l==null?null:new HyperlinkImpl(new SimpleRef<org.zkoss.poi.ss.usermodel.Hyperlink>(l));
	}
	
	public void setSheetName(String name){
		//TODO the syncLevel
		range.setSheetName(name);
	}
	
	public String getSheetName(){
		return getSheet().getSheetName();
	}
	
	public void setSheetOrder(int pos){
		//TODO the syncLevel
		range.setSheetOrder(pos);
	}
	
	public int getSheetOrder(){
		return getBook().getSheetIndex(getSheet());
	}
	
	public void setCellValue(Object value){
		//TODO the syncLevel
		range.setValue(value);
	}
	
	private ModelRef<XBook> getBookRef(){
		return ((BookImpl)getBook()).getRef();
	}
	

	/**
	 * get the first cell style of this range
	 * 
	 * @return cell style if cell is exist, the check row style and column cell style if cell not found, if row and column style is not exist, then return default style of sheet
	 */
	public CellStyle getCellStyle() {
		XSheet sheet = range.getSheet();
		XBook book = sheet.getBook();
		
		int r = range.getRow();
		int c = range.getColumn();
		org.zkoss.poi.ss.usermodel.CellStyle style = null;
		Row row = sheet.getRow(r);
		if (row != null){
			Cell cell = row.getCell(c);
			
			if (cell != null){//cell style
				style = cell.getCellStyle();
			}
			if(style==null && row.isFormatted()){//row sytle
				style = row.getRowStyle();
			}
		}
		if(style==null){//col style
			style = sheet.getColumnStyle(c);
		}
		if(style==null){//default
			style = book.getCellStyleAt((short) 0);
		}
		
		return new CellStyleImpl(getBookRef(), new SimpleRef<org.zkoss.poi.ss.usermodel.CellStyle>(style));		
	}
	
	private ClientAnchor toClientAnchor(SheetAnchor anchor){
		Book book = getSheet().getBook();
		ClientAnchor can = null;
		if(book.getType()==BookType.EXCEL_2003){
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
	
	public Picture addPicture(SheetAnchor anchor,byte[] image,Format format){
		ClientAnchor an = toClientAnchor(anchor);
		org.zkoss.poi.ss.usermodel.Picture pic = range.addPicture(an, image, EnumUtil.toPictureFormat(format));
		return new PictureImpl(getBookRef(), new SimpleRef<org.zkoss.poi.ss.usermodel.Picture>(pic));
	}
	
	public void deletePicture(Picture picture){
		//TODO the syncLevel
		range.deletePicture(((PictureImpl)picture).getNative());
	}
	
	public void movePicture(SheetAnchor anchor,Picture picture){
		//TODO the syncLevel
		ClientAnchor an = toClientAnchor(anchor);
		range.movePicture(((PictureImpl)picture).getNative(), an);
	}
	
	//currently, we only support to modify chart in XSSF
	public Chart addChart(SheetAnchor anchor,ChartData data,Type type, Grouping grouping, LegendPosition pos){
		//TODO the syncLevel
		ClientAnchor an = toClientAnchor(anchor);
		org.zkoss.poi.ss.usermodel.charts.ChartData cdata = ((ChartDataImpl)data).getNative();
		org.zkoss.poi.ss.usermodel.Chart chart = range.addChart(an, cdata, EnumUtil.toChartType(type), EnumUtil.toChartGrouping(grouping), EnumUtil.toLegendPosition(pos));
		return new ChartImpl(getBookRef(), new SimpleRef<org.zkoss.poi.ss.usermodel.Chart>(chart));
	}
	
	//currently, we only support to modify chart in XSSF
	public void deleteChart(Chart chart){
		//TODO the syncLevel
		range.deleteChart(((ChartImpl)chart).getNative());
	}
	
	//currently, we only support to modify chart in XSSF
	public void moveChart(SheetAnchor anchor,Chart chart){
		//TODO the syncLevel
		ClientAnchor an = toClientAnchor(anchor);
		range.moveChart(((ChartImpl)chart).getNative(), an);
	}
	
	
	public Sheet createSheet(String name){
		//TODO the syncLevel
		XBook book = ((BookImpl)getBook()).getNative();
		int n = book.getNumberOfSheets();
		range.createSheet(name);
		
		XSheet sheet = book.getWorksheetAt(n);
		return new SheetImpl(new SimpleRef<XSheet>(sheet));
		
	}
	
	public void deleteSheet(){
		//TODO the syncLevel
		range.deleteSheet();
	}
	
	
	@Override
	public void setColumnWidth(int widthPx) {
		XRange r = range.isWholeColumn()?range:range.getColumns();
		r.setColumnWidth(XUtils.pxToFileChar256(widthPx, ((XBook)range.getSheet().getWorkbook()).getDefaultCharWidth()));
	}
	@Override
	public void setRowHeight(int heightPx) {
		XRange r = range.isWholeRow()?range:range.getRows();
		r.setRowHeight(XUtils.pxToPoint(heightPx));
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
	
	
	public void api4Internal(){
		range.notifyDeleteFriendFocus(null);//by Spreadsheet
		range.notifyMoveFriendFocus(null);//
	}
	
	
	//API of range that n-oone use it.
	
	
	public void apiNoOneUse(){
		
		range.getCount();
		range.getCurrentRegion();
		range.getDependents();
		range.getDirectDependents();
		range.getPrecedents();
		
		range.isCustomHeight();
		
		//range.pasteSpecial(pasteType, pasteOp, SkipBlanks, transpose);		
	}

}
