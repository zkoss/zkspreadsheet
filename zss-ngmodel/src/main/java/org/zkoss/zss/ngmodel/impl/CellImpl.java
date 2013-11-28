package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collection;
import java.util.Date;

import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.util.CellReference;
import org.zkoss.zss.ngmodel.util.Validations;

public class CellImpl extends AbstractCell {
	private static final long serialVersionUID = 1L;
	private AbstractRow row;
	private CellType type = CellType.BLANK;
	private Object value = null;
	private AbstractCellStyle cellStyle;

	transient private ResultWrap formulaResult;// cache

	public CellImpl(AbstractRow row) {
		this.row = row;
	}

	@Override
	public CellType getType() {
		return type;
	}

	@Override
	public boolean isNull() {
		return false;
	}

	@Override
	public int getRowIndex() {
		checkOrphan();
		return row.getIndex();
	}

	@Override
	public int getColumnIndex() {
		checkOrphan();
		return row.getCellIndex(this);
	}

	@Override
	public String getReferenceString() {
		return new CellReference(this).formatAsString();
	}

	@Override
	public void checkOrphan() {
		if (row == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	public NSheet getSheet() {
		checkOrphan();
		return row.getSheet();
	}

	@Override
	public void release() {
		checkOrphan();
		clearValue();
		row = null;
	}

	@Override
	public NCellStyle getCellStyle() {
		return getCellStyle(false);
	}

	@Override
	public NCellStyle getCellStyle(boolean local) {
		if (local || cellStyle != null) {
			return cellStyle;
		}
		checkOrphan();
		cellStyle = (AbstractCellStyle) row.getCellStyle(true);
		AbstractSheet sheet = (AbstractSheet)row.getSheet();
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

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, AbstractCellStyle.class);
		this.cellStyle = (AbstractCellStyle) cellStyle;
	}

	@Override
	protected void evalFormula() {
		if (type == CellType.FORMULA && formulaResult == null) {
			FormulaEngine fe = EngineFactory.getInstance()
					.createFormulaEngine();
			formulaResult = new ResultWrap(fe.evaluate((FormulaExpression) value,
					new FormulaEvaluationContext()));
		}
	}

	@Override
	public CellType getFormulaResultType() {
		checkType(CellType.FORMULA);
		evalFormula();

		return formulaResult.getCellType();
	}

	@Override
	public void clearValue() {
		checkOrphan();
		value = null;
		formulaResult = null;
		if (type == CellType.FORMULA) {
			// clear depends
			Ref ref = new RefImpl(this);
			((AbstractBookSeries) row.getSheet().getBook().getBookSeries())
					.getDependencyTable().clearDependents(ref);
		}
		type = CellType.BLANK;
	}

	@Override
	public void clearFormulaResultCache() {
		formulaResult = null;
	}

	@Override
	public Object getValue(boolean eval) {
		if (eval && type == CellType.FORMULA) {
			evalFormula();
			return formulaResult.getValue();
		}
		return value;
	}

	private boolean isFormula(String string) {
		return string != null && string.startsWith("=") && string.length() > 1;
	}

	@Override
	public void setValue(Object newvalue) {
		if (value != null && value.equals(newvalue)) {
			return;
		}
		clearValue();

		if (newvalue == null) {
			// nothing
		} else if (newvalue instanceof String) {
			if ("".equals(newvalue)) {
				type = CellType.BLANK;
				newvalue = null;
			} else if (isFormula((String) newvalue)) {
				setFormulaValue(((String) newvalue).substring(1));
				return;// break;
			} else {
				type = CellType.STRING;
			}
		} else if (newvalue instanceof FormulaExpression) {
			type = CellType.FORMULA;
		} else if (newvalue instanceof Date) {
			type = CellType.DATE;
		} else if (newvalue instanceof Number) {
			type = CellType.NUMBER;
		} else if (newvalue instanceof ErrorValue) {
			type = CellType.ERROR;
		} else {
			throw new IllegalArgumentException(
					"unsupported type "
							+ newvalue
							+ ", supports NULL, String, Date, Number and Byte(as Error Code)");
		}
		value = newvalue;
	}
	
	private class ResultWrap implements Serializable{
		private static final long serialVersionUID = 1L;
		
		CellType cellType = null;
		Object value = null;
		private ResultWrap(EvaluationResult result){
			Object val = result.getValue();
			ResultType type = result.getType();
			if(type==ResultType.ERROR){
				cellType = CellType.ERROR;
				value = (val instanceof ErrorValue)?(ErrorValue)val:new ErrorValue(ErrorValue.INVALID_VALUE);
			}else if(type==ResultType.SUCCESS){
				setByValue(val);
			}
		}
		
		private void setByValue(Object val){
			if(val==null || "".equals(val)){
				cellType = CellType.BLANK;
				value = null;
			}else if(val instanceof String){
				cellType = CellType.STRING;
				value = (String)val;
			}else if(val instanceof Number){
				cellType = CellType.NUMBER;
				value = (Number)val;
			}else if(val instanceof Date){
				cellType = CellType.NUMBER;
				value = (Date)val;
			}else if(val instanceof Collection){
				//possible a engine return a collection in cell evaluation case? who should take care array formula?
				if(((Collection)val).size()>0){
					setByValue(((Collection)val).iterator().next());
				}else{
					cellType = CellType.BLANK;
					value = null;
				}
			}else if(val.getClass().isArray()){
				//possible a engine return a collection in cell evaluation case? who should take care array formula?
				if(((Object[])val).length>0){
					setByValue(((Object[])val)[0]);
				}else{
					cellType = CellType.BLANK;
					value = null;
				}
			}else{
				cellType = CellType.ERROR;
				value = (val instanceof ErrorValue)?(ErrorValue)val:new ErrorValue(ErrorValue.INVALID_VALUE,"Unknow value type "+val);
			}
			
			
		}
		
		private CellType getCellType(){
			return cellType;
		}
		
		private Object getValue(){
			return value;
		}
	}
}
