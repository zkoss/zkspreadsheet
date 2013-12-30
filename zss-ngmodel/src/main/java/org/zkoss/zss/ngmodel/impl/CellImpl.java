/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/12/01 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ngmodel.impl;

import java.io.Serializable;
import java.util.Collection;
import java.util.Date;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.InvalidateModelOpException;
import org.zkoss.zss.ngmodel.InvalidateModelValueException;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NColumnArray;
import org.zkoss.zss.ngmodel.NComment;
import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.NDataGrid;
import org.zkoss.zss.ngmodel.NHyperlink;
import org.zkoss.zss.ngmodel.NRichText;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaClearContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.util.Validations;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class CellImpl extends AbstractCellAdv {
	private static final long serialVersionUID = 1L;
	private AbstractRowAdv row;
	private int index;
	private NCellValue localValue = null;
	private AbstractCellStyleAdv cellStyle;
	transient private FormulaResultWrap formulaResult;// cache
	
	
	//use another object to reduce object reference size
	private OptFields opts;
	
	private static class OptFields implements Serializable{
		private AbstractHyperlinkAdv hyperlink;
		private AbstractCommentAdv comment;
		private AbstractRichTextAdv richText;
	}

	private OptFields getOpts(boolean create){
		if(opts==null && create){
			opts = new OptFields();
		}
		return opts;
	}

	public CellImpl(AbstractRowAdv row, int index) {
		this.row = row;
		this.index = index;
	}

	@Override
	public CellType getType() {
		NCellValue val = getDataGridValue();
		return val==null?CellType.BLANK:val.getType();
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
		return index;
	}

	@Override
	public String getReferenceString() {
		return new CellRegion(getRowIndex(), getColumnIndex()).getReferenceString();
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
	public void destroy() {
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
		cellStyle = (AbstractCellStyleAdv) row.getCellStyle(true);
		AbstractSheetAdv sheet = (AbstractSheetAdv)row.getSheet();
		if (cellStyle == null) {
			NColumnArray array = sheet.getColumnArray(getColumnIndex());
			if(array!=null){
				cellStyle = (AbstractCellStyleAdv)((AbstractColumnArrayAdv)array).getCellStyle(true);
			}
		}
		if (cellStyle == null) {
			cellStyle = (AbstractCellStyleAdv) sheet.getBook()
					.getDefaultCellStyle();
		}
		return cellStyle;
	}

	@Override
	public void setCellStyle(NCellStyle cellStyle) {
		Validations.argNotNull(cellStyle);
		Validations.argInstance(cellStyle, AbstractCellStyleAdv.class);
		this.cellStyle = (AbstractCellStyleAdv) cellStyle;
	}

	@Override
	protected void evalFormula() {
		if (formulaResult == null) {
			NCellValue val = getDataGridValue();
			if(val!=null &&  val.getType() == CellType.FORMULA){
				FormulaEngine fe = EngineFactory.getInstance()
						.createFormulaEngine();
				formulaResult = new FormulaResultWrap(fe.evaluate((FormulaExpression) val.getValue(),
						new FormulaEvaluationContext(this)));
			}
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
		//clear dg first, it might throw invalidate model operation exception
		validateDataGridValue(null);
		setDataGridValue(null);
		
		clearFormulaDependency();
		clearFormulaResultCache();
		
		OptFields opts = getOpts(false); 
		if(opts!=null){
			opts.richText = null;
		};
	}
	
	
	/*package*/ void clearValueForSet(boolean clearDependency) {
		checkOrphan();
		
		//shouldn't clear type and data grid value for setting, it will set a new value directly later 
		//just clear cell local field.
		//setDataGridValue(null);
		
		//in some situation, we should clear dependency (e.g. old type and new type are both formula)
		if(clearDependency){
			clearFormulaDependency();
		}
		clearFormulaResultCache();
		
		OptFields opts = getOpts(false); 
		if(opts!=null){
			opts.richText = null;
		};
	}

	@Override
	public void clearFormulaResultCache() {
		if(formulaResult!=null){			
			//only clear when there is a formula result, or poi will do full cache scan to clean blank.
			EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		}
		
		formulaResult = null;
	}
	
	@Override
	public boolean isFormulaParsingError() {
		if (getType() == CellType.FORMULA) {
			return ((FormulaExpression)getValue(false)).hasError();
		}
		return false;
	}
	
	private void clearFormulaDependency(){
		Ref ref = new RefImpl(this);
			((AbstractBookSeriesAdv) getSheet().getBook().getBookSeries())
					.getDependencyTable().clearDependents(ref);
	}

	@Override
	public Object getValue(boolean evaluatedVal) {
		NCellValue val = getDataGridValue();
		if (evaluatedVal && val!=null && val.getType() == CellType.FORMULA) {
			evalFormula();
			return this.formulaResult.getValue();
		}
		return val==null?null:val.getValue();
	}

	private boolean isFormula(String string) {
		return string != null && string.startsWith("=") && string.length() > 1;
	}
	
	private NCellValue getDataGridValue(){
		checkOrphan();
		NDataGrid dg = row.getSheet().getDataGrid();
		if(dg!=null){
			return dg.getValue(getRowIndex(),getColumnIndex());
		}
		return localValue;
	}
	

	private void validateDataGridValue(NCellValue value) {
		checkOrphan();
		
		NDataGrid dg = row.getSheet().getDataGrid();
		if(dg!=null && !dg.validateValue(getRowIndex(),getColumnIndex(),value)){
			throw new InvalidateModelValueException("Invalidate Value : "+(value==null?null:value.getValue()) +" at "+getReferenceString());
		}
	}
	
	private void setDataGridValue(NCellValue value){
		checkOrphan();
		NDataGrid dg = row.getSheet().getDataGrid();
		if(dg!=null){
			this.localValue = null;//clear it if it is formula or error previously
			dg.setValue(getRowIndex(),getColumnIndex(),value);
		}else{
			this.localValue = value!=null&&value.getType()==CellType.BLANK?null:value;
		}
	}

	@Override
	public void setValue(Object newVal) {
		NCellValue oldVal = getDataGridValue();
		if( (oldVal==null && newVal==null) ||
			(oldVal != null && oldVal.getValue().equals(newVal))) {
			return;
		}
		
		CellType newType;
		
		if (newVal == null) {
			newType = CellType.BLANK;
		} else if (newVal instanceof String) {
			if ("".equals(newVal)) {
				newType = CellType.BLANK;
				newVal = null;
			} else if (isFormula((String) newVal)) {
				setFormulaValue(((String) newVal).substring(1));
				return;// break;
			} else {
				newType = CellType.STRING;
			}
		} else if (newVal instanceof FormulaExpression) {
			newType = CellType.FORMULA;
		} else if (newVal instanceof Date) {
			newType = CellType.NUMBER;
			newVal = EngineFactory.getInstance().getCalendarUtil().dateToDoubleValue((Date)newVal, false);
		} else if (newVal instanceof Boolean) {
			newType = CellType.BOOLEAN;
		} else if (newVal instanceof Double) {
			newType = CellType.NUMBER;
		} else if (newVal instanceof Number) {
			newType = CellType.NUMBER;
			newVal = ((Number)newVal).doubleValue();
		} else if (newVal instanceof ErrorValue) {
			newType = CellType.ERROR;
		} else {
			throw new IllegalArgumentException(
					"unsupported type "
							+ newVal
							+ ", supports NULL, String, Date, Number and Byte(as Error Code)");
		}

		
		NCellValue newCellVal = new InnerCellValue(newType,newVal);
		validateDataGridValue(newCellVal);
		//should't clear dependency if new type is formula, it clear the dependency already when eval
		clearValueForSet(oldVal!=null && oldVal.getType()==CellType.FORMULA && newType !=CellType.FORMULA);
		
		setDataGridValue(newCellVal);
	}

	private class FormulaResultWrap implements Serializable{
		private static final long serialVersionUID = 1L;
		
		CellType cellType = null;
		Object value = null;
		private FormulaResultWrap(EvaluationResult result){
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
			}else if(val instanceof Boolean){
				cellType = CellType.BOOLEAN;
				value = (Boolean)val;
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

	@Override
	public NHyperlink getHyperlink() {
		OptFields opts = getOpts(false);
		return opts==null?null:opts.hyperlink;
	}

	@Override
	public void setHyperlink(NHyperlink hyperlink) {
		Validations.argInstance(hyperlink, AbstractHyperlinkAdv.class);
		getOpts(true).hyperlink = (AbstractHyperlinkAdv)hyperlink;
	}
	
	@Override
	public NComment getComment() {
		OptFields opts = getOpts(false);
		return opts==null?null:opts.comment;
	}

	@Override
	public void setComment(NComment comment) {
		Validations.argInstance(comment, AbstractCommentAdv.class);
		getOpts(true).comment = (AbstractCommentAdv)comment;
	}

	@Override
	public void setRichText(NRichText text) {
		Validations.argInstance(text, AbstractRichTextAdv.class);
		getOpts(true).richText = (AbstractRichTextAdv)text;
	}

	@Override
	public NRichText getRichText() {
		OptFields opts = getOpts(false);
		return opts==null?null:opts.richText;
	}
	
	@Override
	void setIndex(int newidx) {
		this.index = newidx;
	}
	
	private static class InnerCellValue extends NCellValue{
		private static final long serialVersionUID = 1L;
		private InnerCellValue(CellType type, Object value){
			super(type,value);
		}
	}
}
