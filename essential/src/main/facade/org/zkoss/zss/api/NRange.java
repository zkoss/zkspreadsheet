package org.zkoss.zss.api;

import org.zkoss.poi.ss.usermodel.Cell;
import org.zkoss.poi.ss.usermodel.Row;
import org.zkoss.zss.api.model.NCellStyle;
import org.zkoss.zss.api.model.NSheet;
import org.zkoss.zss.api.model.impl.EnumUtil;
import org.zkoss.zss.model.Range;
import org.zkoss.zss.model.Worksheet;

public class NRange {
	
	enum VisitorLockLevel{
		BOOK,
		NONE//for you just visit and do nothing
	}
	enum BatchLockLevel{
		BOOK,
		SHEET
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
	
	NSheet nsheet;
	
	Range range;
	public NRange(Range range) {
		this.range = range;
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

	
	public boolean isAnyCellProtected(){
		return range.isAnyCellProtected();
	}

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
		if(nsheet!=null){
			return nsheet;
		}
		return nsheet = new NSheet(range.getSheet());
	}

 
	public void clearStyles() {
		range.setStyle(null);//will use default book cell style		
	}

	public void setStyle(NCellStyle nstyle) {
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
	
	
	
	public void batch(NBatchRunner run,BatchLockLevel lock){
		if(lock==BatchLockLevel.BOOK){
			synchronized(range.getSheet().getBook()){
				run.run(this);
			}
		}else if(lock==BatchLockLevel.SHEET){
			synchronized(range.getSheet()){
				run.run(this);
			}
		}
	}
	/**
	 * visit all cells in this range
	 * @param visitor the visitor 
	 * @param create create cell if it doesn't exist, if it is true, it will also lock the sheet
	 * @param lock lock the sheet if you will do any modification of the sheet 
	 */
	public void visit(final NCellVisitor visitor,VisitorLockLevel lock){
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
		
		if(lock!=null && lock!=VisitorLockLevel.NONE){
			if(lock==VisitorLockLevel.BOOK){
				synchronized(range.getSheet().getBook()){
					synchronized(range.getSheet()){
						run.run();
					}
				}
			}else{
				//TODO sheet level?
			}
		}else{
			run.run();
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
		visitor.visit(NRanges.range(getSheet(),r,c));
	}
	
	
}
