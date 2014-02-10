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
import java.util.Date;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.InvalidateModelValueException;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCellStyle;
import org.zkoss.zss.ngmodel.NCellValue;
import org.zkoss.zss.ngmodel.NColumnArray;
import org.zkoss.zss.ngmodel.NComment;
import org.zkoss.zss.ngmodel.NHyperlink;
import org.zkoss.zss.ngmodel.NRichText;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.FormulaClearContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
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
	transient private FormulaResultCellValue formulaResultValue;// cache
	
	
	//use another object to reduce object reference size
	private OptFields opts;
	
	private static class OptFields implements Serializable{
		private AbstractHyperlinkAdv hyperlink;
		private AbstractCommentAdv comment;
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
		NCellValue val = getCellValue();
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
		addCellUpdate();
	}

	@Override
	protected void evalFormula() {
		if (formulaResultValue == null) {
			NCellValue val = getCellValue();
			if(val!=null &&  val.getType() == CellType.FORMULA){
				FormulaEngine fe = EngineFactory.getInstance()
						.createFormulaEngine();
				formulaResultValue = new FormulaResultCellValue(fe.evaluate((FormulaExpression) val.getValue(),
						new FormulaEvaluationContext(this)));
			}
		}
	}

	@Override
	public CellType getFormulaResultType() {
		checkType(CellType.FORMULA);
		evalFormula();

		return formulaResultValue.getCellType();
	}

	@Override
	public void clearValue() {
		checkOrphan();
		setCellValue(null);
		
		clearFormulaDependency();
		clearFormulaResultCache();
		
		OptFields opts = getOpts(false); 
		if(opts!=null){
			// clear for value, don't clear hyperlink
//			opts.hyperlink = null;
		};
		//don't update when sheet is destroying
		if(BookImpl.destroyingSheet.get()!=getSheet()){
			addCellUpdate();
		}
	}
	
	private void addCellUpdate(){
		ModelUpdateUtil.addCellUpdate(getSheet(), getRowIndex(), getColumnIndex());
	}
	
	
	/*package*/ void clearValueForSet(boolean clearDependency) {
		checkOrphan();
		
		//in some situation, we should clear dependency (e.g. old type and new type are both formula)
		if(clearDependency){
			clearFormulaDependency();
		}
		clearFormulaResultCache();
		
		OptFields opts = getOpts(false); 
		if(opts!=null){
			// Clear value only, don't clear hyperlink
//			opts.hyperlink = null;
		};
	}

	@Override
	public void clearFormulaResultCache() {
		if(formulaResultValue!=null){			
			//only clear when there is a formula result, or poi will do full cache scan to clean blank.
			EngineFactory.getInstance().createFormulaEngine().clearCache(new FormulaClearContext(this));
		}
		
		formulaResultValue = null;
	}
	
	@Override
	public boolean isFormulaParsingError() {
		if (getType() == CellType.FORMULA) {
			return ((FormulaExpression)getValue(false)).hasError();
		}
		return false;
	}
	
	/*package*/ void clearFormulaDependency(){
		if(lastRef!=null){
			((AbstractBookSeriesAdv) getSheet().getBook().getBookSeries())
					.getDependencyTable().clearDependents(lastRef);
		}
		lastRef = null;
	}

	@Override
	public Object getValue(boolean evaluatedVal) {
		NCellValue val = getCellValue();
		if (evaluatedVal && val!=null && val.getType() == CellType.FORMULA) {
			evalFormula();
			return this.formulaResultValue.getValue();
		}
		return val==null?null:val.getValue();
	}

	private boolean isFormula(String string) {
		return string != null && string.startsWith("=") && string.length() > 1;
	}
	
	private NCellValue getCellValue(){
		checkOrphan();
		return localValue;
	}
	
	private void setCellValue(NCellValue value){
		checkOrphan();
		this.localValue = value!=null&&value.getType()==CellType.BLANK?null:value;
		
		//clear the dependent's formula result cache
		NBookSeries bookSeries = getSheet().getBook().getBookSeries();
		ModelUpdateUtil.handlePrecedentUpdate(bookSeries,getRef());
	}

	
	private static boolean valueEuqals(Object val1, Object val2){
		return val1==val2||(val1!=null && val1.equals(val2));
	}
	
	@Override
	public void setValue(Object newVal) {
		NCellValue oldVal = getCellValue();
		if( (oldVal==null && newVal==null) ||
			(oldVal != null && valueEuqals(oldVal.getValue(),newVal))) {
			return;
		}
		
		CellType newType;
		
		if (newVal == null) {
			newType = CellType.BLANK;
		} else if (newVal instanceof String) {
			if (isFormula((String) newVal)) {
				setFormulaValue(((String) newVal).substring(1));
				return;// break;
			} else {
				newType = CellType.STRING;
			}
		} else if (newVal instanceof NRichText) {
			newType = CellType.STRING;
		} else if (newVal instanceof FormulaExpression) {
			newType = CellType.FORMULA;
		} else if (newVal instanceof Date) {
			newType = CellType.NUMBER;
			newVal = EngineFactory.getInstance().getCalendarUtil().dateToDoubleValue((Date)newVal);
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
		//should't clear dependency if new type is formula, it clear the dependency already when eval
		clearValueForSet(oldVal!=null && oldVal.getType()==CellType.FORMULA && newType !=CellType.FORMULA);
		
		setCellValue(newCellVal);
		addCellUpdate();
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
		addCellUpdate();
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
		addCellUpdate();
	}
	
	@Override
	void setIndex(int newidx) {
		this.index = newidx;
		rebuildFormulaDependency();
	}
	
	@Override
	void setRow(AbstractRowAdv row){
		checkOrphan();
		this.row = row;
		rebuildFormulaDependency();
	}
	
	private Ref lastRef;

	protected Ref getRef(){
		if(lastRef==null){//keep the last ref for clear dependency (cell is possible be shifted)
			lastRef = new RefImpl(this);
		}
		return lastRef;
	}
	
	private static class InnerCellValue extends NCellValue{
		private static final long serialVersionUID = 1L;
		private InnerCellValue(CellType type, Object value){
			super(type,value);
		}
	}
	
	public String toString(){
		StringBuilder sb = new StringBuilder();
		sb.append("Cell:[").append(getRowIndex()).append(",").append(getColumnIndex()).append("]");
		return sb.toString();
	}
	
	//rebuild formula dependency after shift
	/*package*/ void rebuildFormulaDependency(){
		if(getType()==CellType.FORMULA && lastRef!=null){
			FormulaExpression expr = (FormulaExpression)getValue(false);
			if(!expr.hasError()){
				//clear and rebuild the dependency
				clearFormulaDependency();
				FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
				fe.parse(expr.getFormulaString(), new FormulaParseContext(this ,getRef()));//rebuild the expression to make new dependency with current row,column
			}
		}
		
	}
}
