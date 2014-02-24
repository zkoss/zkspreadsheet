package org.zkoss.zss.ngmodel.impl;

import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NComment;
import org.zkoss.zss.ngmodel.NDataValidation;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.NHyperlink;
/**
 * a help class to hold cell data and apply to another
 * @author Dennis
 */
public class CellBuffer {
	
	boolean nullFlag = true;;
	
	CellType type;
	Object value;
	String formula;
	NCellStyle style;
	
	NComment comment;
	NDataValidation validation;
	NHyperlink hyperlink;
	
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
	public NCellStyle getStyle() {
		return style;
	}
	public void setStyle(NCellStyle style) {
		this.style = style;
	}
	public NComment getComment() {
		return comment;
	}
	public void setComment(NComment comment) {
		this.comment = comment;
	}
	public NDataValidation getValidation() {
		return validation;
	}
	public void setValidation(NDataValidation validation) {
		this.validation = validation;
	}
	public NHyperlink getHyperlink(){
		return hyperlink;
	}
	public void setHyperlink(NHyperlink hyperlink){
		this.hyperlink = hyperlink;
	}
	
	public static CellBuffer bufferAll(NCell cell){
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

	
	public void applyAll(NCell cell){
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

	public void applyValidation(NCell destCell) {
		NDataValidation srcValidation = getValidation();
		//TODO, base on original validation data structure, it it very complicated to past a validation, consider to change the structure.
	}
	
	public void applyStyle(NCell destCell) {
		//style are shared between sheets, could use it directly
		destCell.setCellStyle(getStyle());
	}
	
	public void applyValue(NCell destCell) {
		if(getType()==CellType.FORMULA){
			destCell.setFormulaValue(getFormula());
		}else{
			destCell.setValue(getValue());
		}
	}
	public void applyComment(NCell destCell) {
		NComment comment = getComment();
		destCell.setComment(comment==null?null:((AbstractCommentAdv)comment).clone());
	}
	
	public void applyHyperlink(NCell destCell) {
		NHyperlink link = getHyperlink();
		destCell.setHyperlink(link==null?null:((AbstractHyperlinkAdv)link).clone());
	}
}
