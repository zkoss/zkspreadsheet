package org.zkoss.zss.api;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.poi.hssf.usermodel.HSSFClientAnchor;
import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.CellStyle;
import org.zkoss.poi.ss.usermodel.Chart;
import org.zkoss.poi.ss.usermodel.ClientAnchor;
import org.zkoss.poi.ss.usermodel.Picture;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.usermodel.charts.ChartData;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.poi.xssf.usermodel.XSSFClientAnchor;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NBook.BookType;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NCellStyle.BorderType;
import org.zkoss.zss.api.model.NChart;
import org.zkoss.zss.api.model.NChart.Grouping;
import org.zkoss.zss.api.model.NChart.LegendPosition;
import org.zkoss.zss.api.model.NChart.Type;
import org.zkoss.zss.api.model.NChartData;
import org.zkoss.zss.api.model.NHyperlink.HyperlinkType;
import org.zkoss.zss.api.model.NPicture;
import org.zkoss.zss.api.model.NPicture.Format;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.model.SimpleRef;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;
import org.zkoss.zss.model.impl.DrawingManager;
import org.zkoss.zss.model.impl.SheetCtrl;

/**
 * 1.Range is not handling the protection issue, if you have handle it yourself before calling the api(by calling {@code #isProtected()})
 * @author dennis
 *
 */
public class NRange {
	
	enum SyncLevel{
		BOOK,
		NONE//for you just visit and do nothing
	}
	
	public enum PasteType{
		PASTE_ALL,
		PASTE_ALL_EXCEPT_BORDERS,
		PASTE_COLUMN_WIDTHS,
		PASTE_COMMENTS,
		PASTE_FORMATS/*all formats*/,
		PASTE_FORMULAS/*include values and formulas*/,
		PASTE_FORMULAS_AND_NUMBER_FORMATS,
		PASTE_VALIDATAION,
		PASTE_VALUES,
		PASTE_VALUES_AND_NUMBER_FORMATS;
	}
	
	public enum PasteOperation{
		PASTEOP_ADD,
		PASTEOP_SUB,
		PASTEOP_MUL,
		PASTEOP_DIV,
		PASTEOP_NONE;
	}
	
	public enum ApplyBorderType{
		FULL,
		EDGE_BOTTOM,
		EDGE_RIGHT,
		EDGE_TOP,
		EDGE_LEFT,
		OUTLINE,
		INSIDE,
		INSIDE_HORIZONTAL,
		INSIDE_VERTICAL,
		DIAGONAL,
		DIAGONAL_DOWN,
		DIAGONAL_UP
	}
	
	/** Shift direction of insert and delete**/
	public enum InsertShift{
		DEFAULT,
		RIGHT,
		DOWN
	}
	/** copy origin of insert and delete**/
	public enum InsertCopyOrigin{
		NONE,
		LEFT_ABOVE,
		RIGHT_BELOW,
	}
	/** Shift direction of insert and delete**/
	public enum DeleteShift{
		DEFAULT,
		LEFT,
		UP
	}
	
	public enum SortDataOption{
		TEXT_AS_NUMBERS
	}
	
	public enum AutoFilterOperation{
		AND,
		BOTTOM10,
		BOTOOM10PERCENT,
		OR,
		TOP10,
		TOP10PERCENT,
		VALUES
	}
	
	public enum AutoFillType{
		COPY,
		DAYS,
		DEFAULT,
		FORMATS,
		MONTHS,
		SERIES,
		VALUES,
		WEEKDAYS,
		YEARS,
		GROWTH_TREND,
		LINER_TREND
	}

	
	Range range;
	
	SyncLevel syncLevel = SyncLevel.BOOK;
	
	public void setSyncLevel(SyncLevel syncLevel){
		this.syncLevel = syncLevel;
	}
	
	private SharedContext sharedCtx;
	
	public NRange(NSheet sheet,Range range) {
		this.range = range;
		sharedCtx = new SharedContext(sheet);
	}
	public NRange(Range range,SharedContext ctx) {
		this.range = range;
		sharedCtx = ctx;
	}
	
	
	public NCreator getCreator(){
		return new NCreator(this);
	}
	
	public Range getNative(){
		return range;
	}
	
	static class SharedContext{
		NSheet nsheet;
		List<MergeArea> mergeAreas;
		
		private SharedContext(NSheet nsheet){
			this.nsheet = nsheet;
			
		}
		
		public NSheet getSheet(){
			return nsheet;
		}
		
		
		public List<MergeArea> getMergeAreas(){
			initMergeRangesCache();
			return mergeAreas;
		}
		
		
		private void initMergeRangesCache(){
			if(mergeAreas==null){//TODO a better way to cache(index) this.(I this MergeMatrixHelper is not good enough now)
				Worksheet sheet = nsheet.getNative();
				int sz = sheet.getNumMergedRegions();
				mergeAreas = new ArrayList<MergeArea>(sz);
				for(int j = sz - 1; j >= 0; --j) {
					final CellRangeAddress addr = sheet.getMergedRegion(j);
					mergeAreas.add(new MergeArea(addr.getFirstColumn(), addr.getFirstRow(), addr.getLastColumn(), addr.getLastRow()));
				}
			}
		}
		
		public void resetMergeAreas(){
			mergeAreas = null;
		}
	}
	private static class MergeArea{
		int r,lr,c,lc;
		public MergeArea(int c,int r,int lc,int lr){
			this.c = c;
			this.r = r;
			this.lc = lc;
			this.lr = lr;
		}

		public boolean contains(int c, int r) {
			return c >= this.c && c <= this.lc && r >= this.r && r <= this.lr;
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
		NRange other = (NRange) obj;
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
	public boolean paste(NRange dest) {		
		return pasteSpecial(dest,PasteType.PASTE_ALL,PasteOperation.PASTEOP_NONE,false,false);
	}
	
	/**
	 * @param dest the destination 
	 * @param transpose TODO
	 * @return true if paste successfully, past to a protected sheet with any
	 *         locked cell in the destination range will always cause past fail.
	 */
	public boolean pasteSpecial(NRange dest,PasteType type,PasteOperation op,boolean skipBlanks,boolean transpose) {
//		if(!isAnyCellProtected()){ // ranges seems this in copy/paste already
		//TODO the syncLevel
		Range r = range.pasteSpecial(dest.getNative(), EnumUtil.toRangePasteTypeNative(type), EnumUtil.toRangePasteOpNative(op), skipBlanks, transpose);
		return r!=null;
//		}
	}


	public void clearContents() {
		//TODO the syncLevel
		range.clearContents();		
	}
	
	public NSheet getSheet(){
		return sharedCtx.getSheet();
	}

 
	public void clearStyles() {
		setStyle(null);//will use default book cell style		
	}

	public void setStyle(final NCellStyle nstyle) {
		//TODO the syncLevel
		range.setStyle(nstyle==null?null:nstyle.getNative());
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
	
	public void sync(NRangeSyncRunner run){
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
	public void visit(final NRangeCellVisitor visitor){
		visit0(visitor,syncLevel);
	}
	private void visit0(final NRangeCellVisitor visitor,SyncLevel sync){
		final int r=getRow();
		final int lr=getLastRow();
		final int c=getColumn();
		final int lc=getLastColumn();
		
		Runnable run = new Runnable(){
			public void run(){
				for(int i=r;i<=lr;i++){
					for(int j=c;j<=lc;j++){
						visitCell(visitor,i,j);
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
	
	private void visitCell(NRangeCellVisitor visitor,int r, int c){
		boolean create = visitor.createIfNotExist(r,c);
		Worksheet sheet = range.getSheet();
		Row row = sheet.getRow(r);
		if(row==null){
			if(create){
				row = sheet.createRow(r);
			}else{
				return;
			}
		}
		Cell cell = row.getCell(c);
		if(cell==null){
			if(create){
				cell = row.createCell(c);
			}else{
				return;
			}
		}
		visitor.visit(new NRange(Ranges.range(range.getSheet(),r,c),sharedCtx));
	}

	public NBook getBook() {
		return getSheet().getBook();
	}
	
	public void applyBordersAround(BorderType borderType,String htmlColor){
		applyBorders(ApplyBorderType.OUTLINE,borderType, htmlColor);
	}
	
	public void applyBorders(ApplyBorderType type,BorderType borderType,String htmlColor){
		//TODO the syncLevel
		range.setBorders(EnumUtil.toRangeApplyBorderType(type), EnumUtil.toRangeBorderType(borderType), htmlColor);
	}

	
	public boolean hasMergeCell(){
		final Result<Boolean> result = new Result<Boolean>(Boolean.FALSE);
		visit0(new NRangeCellVisitor(){
			public boolean createIfNotExist(int row, int column) {
				return false;
			}
			@Override
			public void visit(NRange cellRange) {
				int c = cellRange.getColumn();
				int r = cellRange.getRow();
				for(MergeArea ma:sharedCtx.getMergeAreas()){//TODO index, not loop
					if(ma.contains(c, r)){
						result.set(Boolean.TRUE);
						return;
					}
				}
			}},SyncLevel.NONE);
		return result.get();
	}
	
	
	static public class Result<T> {
		T r;
		public Result(){}
		public Result(T r){
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

	
	public NRange getShiftedRange(int rowOffset,int colOffset){
		NRange offsetRange = new NRange(range.getOffset(rowOffset, colOffset),sharedCtx);
		return offsetRange;
	}
	
	
	public NRange getCellRange(int rowOffset,int colOffset){
		NRange cellRange = new NRange(Ranges.range(range.getSheet(),getRow()+rowOffset,getColumn()+colOffset),sharedCtx);
		return cellRange;
	}
	
	/** get the top-left cell range of this range**/
	public NRange getLeftTop() {
		return getCellRange(0,0);
	}
	
	/**
	 *  Return a range that represents all columns and between the first-row and last-row of this range
	 **/
	public NRange getRowRange(){
		return new NRange(range.getRows(),sharedCtx);
	}
	
	/**
	 *  Return a range that represents all rows and between the first-column and last-column of this range
	 **/
	public NRange getColumnRange(){
		return new NRange(range.getColumns(),sharedCtx);
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
		NRange index = null;
		int r = getRow();
		int c = getColumn();
		int lr = getLastRow();
		int lc = getLastColumn();
		
		index = NRanges.range(this.getSheet(),r,c,sortByRows?r:lr,sortByRows?lc:c);
		
		sort(index,desc,header,matchCase,sortByRows,dataOption,
				null,false,null,null,false,null);
	}
	
	public void sort(NRange index1,
			boolean desc1,
			boolean header, 
			/*int orderCustom, //not implement*/
			boolean matchCase, 
			boolean sortByRows, 
			/*int sortMethod, //not implement*/
			SortDataOption dataOption1,
			NRange index2,boolean desc2,SortDataOption dataOption2,
			NRange index3,boolean desc3,SortDataOption dataOption3){
		
		//TODO the syncLevel
		
		//TODO review the full impl for range1,range2,range3
		
		range.sort(index1==null?null:index1.getNative(), desc1, 
				index2==null?null:index2.getNative()/*rng2*/, -1 /*type*/, desc2/*desc2*/, 
				index3==null?null:index3.getNative()/*rng3*/, desc3/*desc3*/,
				header?BookHelper.SORT_HEADER_YES:BookHelper.SORT_HEADER_NO/*header*/,
				-1/*orderCustom*/, matchCase, sortByRows, -1/*sortMethod*/, 
				dataOption1==null?-1:EnumUtil.toRangeSortDataOption(dataOption1)/*dataOption1*/,
				dataOption2==null?-1:EnumUtil.toRangeSortDataOption(dataOption2)/*dataOption2*/,
				dataOption3==null?-1:EnumUtil.toRangeSortDataOption(dataOption3)/*dataOption3*/);
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
	
	public void fill(NRange dest,AutoFillType fillType){
		//TODO the syncLevel
		range.autoFill(dest.getNative(), EnumUtil.toRangeAutoFillType(fillType));
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
	
	public String getEditText(){
		return range.getEditText();
	}
	
	public void setEditText(String editText){
		//TODO the syncLevel
		range.setEditText(editText);
	}
	
	
	//TODO need to verify the object type
	public Object getValue(){
		return range.getValue();
	}
	
	public void eanbleDisplayGridlines(boolean enable){
		//TODO the syncLevel
		range.setDisplayGridlines(enable);
	}
	
	public boolean isDisplayGridlines(){
		return getSheet().isDisplayGridlines();
	}
	
	public void setHidden(boolean hidden){
		//TODO the syncLevel
		range.setHidden(hidden);
	}
	
	public void setHyperlink(HyperlinkType type,String address,String displayLabel){
		//TODO the syncLevel
		range.setHyperlink(EnumUtil.toHyperlinkType(type), address, displayLabel);
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
	
	public void setValue(Object value){
		//TODO the syncLevel
		range.setValue(value);
	}
	
	

	/**
	 * get the first cell style of this range
	 * 
	 * @return cell style if cell is exist, the check row style and column cell style if cell not found, if row and column style is not exist, then return default style of sheet
	 */
	public NCellStyle getCellStyle() {
		Worksheet sheet = range.getSheet();
		Book book = sheet.getBook();
		
		int r = range.getRow();
		int c = range.getColumn();
		CellStyle style = null;
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
		
		return new NCellStyle(getBook().getRef(), new SimpleRef<CellStyle>(style));		
	}
	
	private ClientAnchor toClientAnchor(NSheetAnchor anchor){
		NBook book = getSheet().getBook();
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
	
	public NPicture addPicture(NSheetAnchor anchor,byte[] image,Format format){
		ClientAnchor an = toClientAnchor(anchor);
		Picture pic = range.addPicture(an, image, EnumUtil.toPictureFormat(format));
		return new NPicture(getBook().getRef(), new SimpleRef<Picture>(pic));
	}
	
	public List<NPicture> getPictures(){
		//TODO the syncLevel
		NBook book = getSheet().getBook();
		DrawingManager dm = ((SheetCtrl)getSheet().getBook().getNative()).getDrawingManager();
		List<NPicture> pictures = new ArrayList<NPicture>();
		for(Picture pic:dm.getPictures()){
			pictures.add(new NPicture(book.getRef(), new SimpleRef<Picture>(pic)));
		}
		return pictures;
	}
	
	public void deletePicture(NPicture picture){
		//TODO the syncLevel
		range.deletePicture(picture.getNative());
	}
	
	public void movePicture(NSheetAnchor anchor,NPicture picture){
		//TODO the syncLevel
		ClientAnchor an = toClientAnchor(anchor);
		range.movePicture(picture.getNative(), an);
	}
	
	//currently, we only support to modify chart in XSSF
	public NChart addChart(NSheetAnchor anchor,NChartData data,Type type, Grouping grouping, LegendPosition pos){
		//TODO the syncLevel
		ClientAnchor an = toClientAnchor(anchor);
		ChartData cdata = data.getNative();
		Chart chart = range.addChart(an, cdata, EnumUtil.toChartType(type), EnumUtil.toChartGrouping(grouping), EnumUtil.toLegendPosition(pos));
		return new NChart(getBook().getRef(), new SimpleRef<Chart>(chart));
	}
	
	public List<NChart> getCharts(){
		//TODO the syncLevel
		NBook book = getSheet().getBook();
		DrawingManager dm = ((SheetCtrl)getSheet().getBook().getNative()).getDrawingManager();
		List<NChart> charts = new ArrayList<NChart>();
		for(Chart chart:dm.getCharts()){
			charts.add(new NChart(book.getRef(), new SimpleRef<Chart>(chart)));
		}
		return charts;
	}
	
	//currently, we only support to modify chart in XSSF
	public void deleteChart(NChart chart){
		//TODO the syncLevel
		range.deleteChart(chart.getNative());
	}
	
	//currently, we only support to modify chart in XSSF
	public void moveChart(NSheetAnchor anchor,NChart chart){
		//TODO the syncLevel
		ClientAnchor an = toClientAnchor(anchor);
		range.moveChart(chart.getNative(), an);
	}
	
	
	public NSheet createSheet(String name){
		//TODO the syncLevel
		Book book = getSheet().getBook().getNative();
		int n = book.getNumberOfSheets();
		range.createSheet(name);
		
		Worksheet sheet = book.getWorksheetAt(n);
		return new NSheet(new SimpleRef<Worksheet>(sheet));
		
	}
	
	public void deleteSheet(){
		//TODO the syncLevel
		range.deleteSheet();
	}
	
	
	//api that need special object wrap
	
	
	public void apiSpecialWrapObject(){

		range.getFormatText();//FormatText
		range.getHyperlink();//Hyperlink
		
		range.getRichEditText();//RichTextString
		range.getText();//RichTextString (what is the difference of getRichEditText)
		
		
		range.validate("");//DataValidation
		
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
