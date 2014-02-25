package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCellStyle;
import org.zkoss.zss.model.SComment;
import org.zkoss.zss.model.SDataValidation;
import org.zkoss.zss.model.SHyperlink;
import org.zkoss.zss.model.SCell.CellType;
/**
 * a help class to hold cell data and apply to another
 * @author Dennis
 */
public class CellBuffer {
	
	boolean nullFlag = true;;
	
	CellType type;
	Object value;
	String formula;
	SCellStyle style;
	
	SComment comment;
	SDataValidation validation;
	SHyperlink hyperlink;
	
	public CellBuffer(){
	}
	
	public boolean isNull(){
		return nullFlag;
	}
	
	public void setNull(boolean nullFlag){
		this.nullFlag = nullFlag;
	}
	
	public CellType getType() {
		return type;
	}
	public void setType(CellType type) {
		this.type = type;
	}
	public Object getValue() {
		return value;
	}
	public void setValue(Object value) {
		this.value = value;
	}
	public String getFormula() {
		return formula;
	}
	public void setFormula(String formula) {
		this.formula = formula;
	}
	public SCellStyle getStyle() {
		return style;
	}
	public void setStyle(SCellStyle style) {
		this.style = style;
	}
	public SComment getComment() {
		return comment;
	}
	public void setComment(SComment comment) {
		this.comment = comment;
	}
	public SDataValidation getValidation() {
		return validation;
	}
	public void setValidation(SDataValidation validation) {
		this.validation = validation;
	}
	public SHyperlink getHyperlink(){
		return hyperlink;
	}
	public void setHyperlink(SHyperlink hyperlink){
		this.hyperlink = hyperlink;
	}
	
	public static CellBuffer bufferAll(SCell cell){
		CellBuffer buffer = new CellBuffer();
		if(!cell.isNull()){
			buffer.setNull(false);
			buffer.setType(cell.getType());
			
			if(cell.getType() == CellType.FORMULA){
				buffer.setFormula(cell.getFormulaValue());
			}else{
				buffer.setValue(cell.getValue());
			}
			
			buffer.setStyle(cell.getCellStyle());
			buffer.setHyperlink(cell.getHyperlink());
			buffer.setComment(cell.getComment());
			buffer.setValidation(cell.getSheet().getDataValidation(cell.getRowIndex(), cell.getColumnIndex()));
		}
		return buffer;
	}

	
	public void applyAll(SCell cell){
		if(isNull()){
			cell.getSheet().clearCell(cell.getRowIndex(), cell.getColumnIndex(), cell.getRowIndex(), cell.getColumnIndex());
		}else{
			applyValue(cell);
			applyStyle(cell);
			applyHyperlink(cell);
			applyComment(cell);
			applyValidation(cell);
		}
	}

	public void applyValidation(SCell destCell) {
		SDataValidation srcValidation = getValidation();
		//TODO, base on original validation data structure, it it very complicated to past a validation, consider to change the structure.
	}
	
	public void applyStyle(SCell destCell) {
		//style are shared between sheets, could use it directly
		destCell.setCellStyle(getStyle());
	}
	
	public void applyValue(SCell destCell) {
		if(getType()==CellType.FORMULA){
			destCell.setFormulaValue(getFormula());
		}else{
			destCell.setValue(getValue());
		}
	}
	public void applyComment(SCell destCell) {
		SComment comment = getComment();
		destCell.setComment(comment==null?null:((AbstractCommentAdv)comment).clone());
	}
	
	public void applyHyperlink(SCell destCell) {
		SHyperlink link = getHyperlink();
		destCell.setHyperlink(link==null?null:((AbstractHyperlinkAdv)link).clone());
	}
}
