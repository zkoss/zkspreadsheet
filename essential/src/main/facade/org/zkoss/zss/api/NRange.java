package org.zkoss.zss.api;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NCellStyle.BorderType;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;
import org.zkoss.zss.model.impl.BookHelper;

/**
 * 1.Range is not handling the protection issue, if you have handle it yourself before calling the api(by calling {@code #isProtected()})
 * @author dennis
 *
 */
public class NRange {
	
	enum LockLevel{
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
	
	Range range;
	
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
	public NGetter getGetter(){
		return new NGetter(this);
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
		Range r = range.pasteSpecial(dest.getNative(), EnumUtil.toRangePasteTypeNative(type), EnumUtil.toRangePasteOpNative(op), skipBlanks, transpose);
		return r!=null;
//		}
	}


	public void clearContents() {
		range.clearContents();		
	}
	
	public NSheet getSheet(){
		return sharedCtx.getSheet();
	}

 
	public void clearStyles() {
		range.setStyle(null);//will use default book cell style		
	}

	public void setStyle(final NCellStyle nstyle) {
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
	
//	public void batch(NBatchRunner run){
//		batch(run,LockLevel.BOOK);
//	}
	public void batch(NRangeBatchRunner run,LockLevel lock){
		switch(lock){
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
//	public void visit(NCellVisitor visitor){
//		visit(visitor,LockLevel.BOOK);
//	}
	/**
	 * visit all cells in this range, make sure you call this in a limited range, 
	 * don't use it for all row/column selection, it will spend much time to iterate the cell 
	 * @param visitor the visitor 
	 * @param create create cell if it doesn't exist, if it is true, it will also lock the sheet
	 * @param lock lock the sheet if you will do any modification of the sheet 
	 */
	public void visit(final NRangeCellVisitor visitor,LockLevel lock){
		final int r=getRow();
		final int lr=getLastRow();
		final int c=getColumn();
		final int lc=getLastColumn();
		
		Runnable run = new Runnable(){
			public void run(){
				for(int i=r;i<=lr;i++){
					for(int j=c;j<=lc;j++){
						visit0(visitor,i,j);
					}
				}
			}
		};
		
		switch(lock){
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
	
	private void visit0(NRangeCellVisitor visitor,int r, int c){
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
	
	public void applyBorderAround(BorderType borderType,String htmlColor){
		range.borderAround(EnumUtil.toRangeBorderType(borderType), htmlColor);
	}
	
	public void applyBorder(ApplyBorderType type,BorderType borderType,String htmlColor){
		range.setBorders(EnumUtil.toRangeApplyBorderType(type), EnumUtil.toRangeBorderType(borderType), htmlColor);
	}

	
	public boolean hasMergeCell(){
		final Result<Boolean> result = new Result<Boolean>(Boolean.FALSE);
		visit(new NRangeCellVisitor(){
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
			}}
		,LockLevel.NONE);
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
		range.merge(across);
	}
	
	public void unMerge(){
		range.unMerge();
	}

	
	public NRange getShiftRange(int rowOffset,int colOffset){
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
	
	/** **/
	public NRange getRowRange(){
		return new NRange(range.getRows(),sharedCtx);
	}
	
	public NRange getColumnRange(){
		return new NRange(range.getColumns(),sharedCtx);
	}
	
	public boolean isContainWholeRow(){
		//TODO, the impl of Ref/Range is opposite to my concept , have to check this
		//original range, wholeColumn means the 'a column' if full selectced, which means, it's rows are full selected. 
		return range.isWholeColumn();
	}
	public boolean isContainWholeColumn(){
		//TODO, the impl of Ref/Range is opposite to my concept , have to check this
		return range.isWholeRow();
	}
	public boolean isContainWholeSheet(){
		return range.isWholeSheet();
	}
	
	public void insert(InsertShift shift,InsertCopyOrigin copyOrigin){
		range.insert(EnumUtil.toRangeInsertShift(shift), EnumUtil.toRangeInsertCopyOrigin(copyOrigin));
	}
	
	public void delete(DeleteShift shift){
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
}
