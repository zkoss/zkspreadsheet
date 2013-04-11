package org.zkoss.zss.api;

import org.zkoss.zss.model.Range;

public class NRange {
	
	enum PasteType{
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
	
	enum PasteOperation{
		PASTEOP_ADD,
		PASTEOP_SUB,
		PASTEOP_MUL,
		PASTEOP_DIV,
		PASTEOP_NONE;
	}
	
	
	Range range;
	public NRange(Range range) {
		this.range = range;
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
		Range r = range.pasteSpecial(dest.getNative(), toPasteTypeNative(type), toPasteOpNative(op), skipBlanks, transpose);
		return r!=null;
//		}
	}
	
	private <T> void assertArgNotNull(T obj,String name){
		if(obj == null){
			throw new IllegalArgumentException("argument "+name==null?"":name+" is null");
		}
	}


	private int toPasteOpNative(PasteOperation op) {
		assertArgNotNull(op,"paste operation");
		switch(op){
		case PASTEOP_ADD:
			return Range.PASTEOP_ADD;
		case PASTEOP_SUB:
			return Range.PASTEOP_SUB;
		case PASTEOP_MUL:
			return Range.PASTEOP_MUL;
		case PASTEOP_DIV:
			return Range.PASTEOP_DIV;
		case PASTEOP_NONE:
			return Range.PASTEOP_NONE;
		}
		throw new IllegalArgumentException("unknow paste operation "+op);
	}


	private int toPasteTypeNative(PasteType type) {
		assertArgNotNull(type,"paste type");
		switch(type){
		case PASTE_ALL:
			return Range.PASTE_ALL;
		case PASTE_ALL_EXCEPT_BORDERS:
			return Range.PASTE_ALL_EXCEPT_BORDERS;
		case PASTE_COLUMN_WIDTHS:
			return Range.PASTE_COLUMN_WIDTHS;
		case PASTE_COMMENTS:
			return Range.PASTE_COMMENTS;
		case PASTE_FORMATS:
			return Range.PASTE_FORMATS;
		case PASTE_FORMULAS:
			return Range.PASTE_FORMULAS;
		case PASTE_FORMULAS_AND_NUMBER_FORMATS:
			return Range.PASTE_FORMULAS_AND_NUMBER_FORMATS;
		case PASTE_VALIDATAION:
			return Range.PASTE_VALIDATAION;
		case PASTE_VALUES:
			return Range.PASTE_VALUES;
		case PASTE_VALUES_AND_NUMBER_FORMATS:
			return Range.PASTE_VALUES_AND_NUMBER_FORMATS;
		}
		throw new IllegalArgumentException("unknow paste operation "+type);
	}


	public void clearContents() {
		range.clearContents();		
	}


	public void clearStyles() {
		range.setStyle(null);//will use default book cell style
	}
}
