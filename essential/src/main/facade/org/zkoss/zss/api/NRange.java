package org.zkoss.zss.api;

import java.util.ArrayList;
import java.util.List;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.poi.ss.util.CellRangeAddress;
import org.zkoss.zss.api.model.NBook;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Ranges;
import org.zkoss.zss.model.Worksheet;

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
	
	public enum ApplyBorderLineStyle{
	    NONE,
	    THIN,
	    MEDIUM,
	    DASHED,
	    HAIR,
	    THICK,
	    DOUBLE,
	    DOTTED,
	    MEDIUM_DASHED,
	    DASH_DOT,
	    MEDIUM_DASH_DOT,
	    DASH_DOT_DOT,
	    MEDIUM_DASH_DOT_DOT,
	    SLANTED_DASH_DOT;
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
		return new NCreator(range);
	}
	public NGetter getGetter(){
		return new NGetter(range);
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
	public void batch(NBatchRunner run,LockLevel lock){
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
	public void visit(final NCellVisitor visitor,LockLevel lock){
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
	
	private void visit0(NCellVisitor visitor,int r, int c){
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
	
	public void applyBorderAround(ApplyBorderLineStyle lineStyle,String htmlColor){
		range.borderAround(EnumUtil.toCellBorderLineStyle(lineStyle), htmlColor);
	}
	
	public void applyBorder(ApplyBorderType type,ApplyBorderLineStyle lineStyle,String htmlColor){
		range.setBorders(EnumUtil.toCellApplyBorderType(type), EnumUtil.toCellBorderLineStyle(lineStyle), htmlColor);
	}

	
	public boolean hasMergeCell(){
		final Result<Boolean> result = new Result<Boolean>(Boolean.FALSE);
		visit(new NCellVisitor(){
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

	
	public NRange getCellRange(int rowOffset,int colOffset){
		NRange cellRange = new NRange(Ranges.range(range.getSheet(),getRow()+rowOffset,getColumn()+colOffset),sharedCtx);
		return cellRange;
	}
	
	/** get the top-left cell range of this range**/
	public NRange getFirst() {
		return getCellRange(0,0);
	}
}
