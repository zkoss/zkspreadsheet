package org.zkoss.zss.ngmodel;

import org.zkoss.zss.ngmodel.util.Validations;


public class PasteOption {

	public enum PasteType{
		ALL,/*BookHelper.INNERPASTE_FORMATS + BookHelper.INNERPASTE_VALUES_AND_FORMULAS + BookHelper.INNERPASTE_COMMENTS + BookHelper.INNERPASTE_VALIDATION;*/ 
		ALL_EXCEPT_BORDERS,/*PASTE_ALL - BookHelper.INNERPASTE_BORDERS;*/
		COLUMN_WIDTHS,/* = BookHelper.INNERPASTE_COLUMN_WIDTHS;*/
		COMMENTS,/* = BookHelper.INNERPASTE_COMMENTS;*/
		FORMATS,/* = BookHelper.INNERPASTE_FORMATS; //all formats*/
		FORMULAS,/* = BookHelper.INNERPASTE_VALUES_AND_FORMULAS; //include values and formulas*/
		FORMULAS_AND_NUMBER_FORMATS,/* = PASTE_FORMULAS + BookHelper.INNERPASTE_NUMBER_FORMATS;*/
		VALIDATAION,/* = BookHelper.INNERPASTE_VALIDATION;*/
		VALUES,/* = BookHelper.INNERPASTE_VALUES;*/
		VALUES_AND_NUMBER_FORMATS/* = PASTE_VALUES + BookHelper.INNERPASTE_NUMBER_FORMATS;*/
	}
	
	public enum PasteOperation{
		ADD,/* = BookHelper.PASTEOP_ADD;*/
		SUB,/* = BookHelper.PASTEOP_SUB;*/
		MUL,/* = BookHelper.PASTEOP_MUL;*/
		DIV,/* = BookHelper.PASTEOP_DIV;*/
		NONE/* = BookHelper.PASTEOP_NONE;*/
	}
	
	boolean skipBlank = false;
//	boolean transport = false;
	
	PasteType pasteType = PasteType.ALL;
	PasteOperation pasteOperation = PasteOperation.NONE;
	
	public boolean isSkipBlank() {
		return skipBlank;
	}
	public void setSkipBlank(boolean skipBlank) {
		this.skipBlank = skipBlank;
	}
	public PasteType getPasteType() {
		return pasteType;
	}
	public void setPasteType(PasteType pasteType) {
		Validations.argNotNull(pasteType);
		this.pasteType = pasteType;
	}
	public PasteOperation getPasteOperation() {
		return pasteOperation;
	}
	public void setPasteOperation(PasteOperation pasteOperation) {
		Validations.argNotNull(pasteOperation);
		this.pasteOperation = pasteOperation;
	}
}
