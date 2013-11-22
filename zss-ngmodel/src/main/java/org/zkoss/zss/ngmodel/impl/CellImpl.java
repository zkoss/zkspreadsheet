package org.zkoss.zss.ngmodel.impl;

import java.util.Date;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

public class CellImpl extends AbstractCell {

	private AbstractRow row;
	private CellType type = CellType.BLANK;
	private Object value = null;
	private AbstractCellStyle cellStyle;
	
	private EvaluationResult formulaResult;

	public CellImpl(AbstractRow row) {
		this.row = row;
	}

	public CellType getType() {
		return type;
	}

	public boolean isNull() {
		return false;
	}

	public int getRowIndex() {
		checkOrphan();
		return row.getIndex();
	}

	public int getColumnIndex() {
		checkOrphan();
		return row.getCellIndex(this);
	}

	public String asString(boolean enableSheetName) {
		return new CellReference(enableSheetName ? row.getSheet()
				.getSheetName() : null, this).formatAsString();
	}

	protected void checkOrphan() {
		if (row == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	void release() {
		row = null;
	}

	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	public NCellStyle getCellStyle(boolean local) {
		if (local || cellStyle != null) {
			return cellStyle;
		}
		checkOrphan();
		cellStyle = (AbstractCellStyle) row.getCellStyle(true);
		AbstractSheet sheet = row.getSheet();
		if (cellStyle == null) {
			cellStyle = (AbstractCellStyle) sheet.getColumn(getColumnIndex())
					.getCellStyle(true);
		}
		if (cellStyle == null) {
			cellStyle = (AbstractCellStyle) sheet.getBook()
					.getDefaultCellStyle();
		}
		return cellStyle;
	}

	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, AbstractCellStyle.class);
		this.cellStyle = (AbstractCellStyle) cellStyle;
	}

	@Override
	protected void evalFormula(){
		if(type==CellType.FORMULA && formulaResult==null){
			//TODO evaluate
			FormulaEngine fe = EngineFactory.getInstance().getFormulaEngine();
			formulaResult = fe.evaluate((FormulaExpression)value,new FormulaEvaluationContext());
		}
	}
	
	public CellType getFormulaResultType() {
		checkType(CellType.FORMULA);
		evalFormula();
		return formulaResult.getType();
	}

	public void clearValue() {
		value = null;
		formulaResult = null;
		type = CellType.BLANK;
	}

	public void clearFormulaResultCache() {
		formulaResult = null;
	}
	
	public Object getValue(boolean eval){
		if(eval && type==CellType.FORMULA){
			evalFormula();
			return formulaResult.getValue();
		}
		return value;
	}
	
	private boolean isFormula(String string){
		return string!=null && string.startsWith("=") && string.length()>1;
	}
	
	public  void setValue(Object newvalue){
		if(value!=null && value.equals(newvalue)){
			return;
		}
		clearValue();
		
		if(newvalue==null){
			//nothing
		}else if(newvalue instanceof String){
			if("".equals(newvalue)){
				type = CellType.BLANK;
				newvalue = null;
			}else if(isFormula((String)newvalue)){
				setFormulaValue(((String)newvalue).substring(1));
				return;//break;
			}else{
				type = CellType.STRING;
			}
		}else if(newvalue instanceof FormulaExpression){
			type = CellType.FORMULA;
		}else if(newvalue instanceof Date){
			type = CellType.DATE;
		}else if(newvalue instanceof Number){
			type = CellType.NUMBER;
		}else if(newvalue instanceof ErrorValue){
			type = CellType.ERROR;
		}else{
			throw new IllegalArgumentException("unsupported type "+newvalue + ", supports NULL, String, Date, Number and Byte(as Error Code)");
		}
		value = newvalue;
	}
}
