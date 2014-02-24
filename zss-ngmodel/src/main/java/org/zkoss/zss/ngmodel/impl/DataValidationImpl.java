package org.zkoss.zss.ngmodel.impl;

import java.util.Collections;
import java.util.LinkedList;
import java.util.List;

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.ErrorValue;
import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NBookSeries;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.util.Validations;

public class DataValidationImpl extends AbstractDataValidationAdv {

	private static final long serialVersionUID = 1L;
	AbstractSheetAdv sheet;
	final private String id;
	
	private ErrorStyle errorStyle = ErrorStyle.STOP;//default stop
	private boolean emptyCellAllowed = true;//default true
	private boolean showDropDownArrow;
	private boolean showPromptBox;
	private boolean showErrorBox;
	private String promptBoxTitle;
	private String promptBoxText;
	private String errorBoxTitle;
	private String errorBoxText;
	private CellRegion region;
	private ValidationType validationType = ValidationType.ANY;
	private OperatorType operatorType = OperatorType.BETWEEN;
	
	
	private FormulaExpression value1Expr;
	private FormulaExpression value2Expr;
	private Object evalValue1Result;
	private Object evalValue2Result;
	
	private boolean evaluated = false;
	
	public DataValidationImpl(AbstractSheetAdv sheet,String id){
		this.sheet = sheet;
		this.id = id;
	}
	
	public String getId(){
		return id;
	}
	
	public NSheet getSheet(){
		return sheet;
	}
	
	@Override
	public void checkOrphan() {
		if (sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	public void destroy() {
		checkOrphan();
		clearFormulaDependency();
		clearFormulaResultCache();
		sheet = null;
	}
	
	@Override
	public ErrorStyle getErrorStyle() {
		return errorStyle;
	}

	@Override
	public void setErrorStyle(ErrorStyle errorStyle) {
		Validations.argNotNull(errorStyle);
		this.errorStyle = errorStyle;
	}

	@Override
	public void setEmptyCellAllowed(boolean allowed) {
		this.emptyCellAllowed = allowed;
	}

	@Override
	public boolean isEmptyCellAllowed() {
		return emptyCellAllowed;
	}

	@Override
	public void setShowDropDownArrow(boolean show) {
		showDropDownArrow = show;
	}

	@Override
	public boolean isShowDropDownArrow() {
		return showDropDownArrow;
	}

	@Override
	public void setShowPromptBox(boolean show) {
		showPromptBox = show;
	}

	@Override
	public boolean isShowPromptBox() {
		return showPromptBox;
	}

	@Override
	public void setShowErrorBox(boolean show) {
		showErrorBox = show;
	}

	@Override
	public boolean isShowErrorBox() {
		return showErrorBox;
	}

	@Override
	public void setPromptBox(String title, String text) {
		promptBoxTitle = title;
		promptBoxText = text;
	}

	@Override
	public String getPromptBoxTitle() {
		return promptBoxTitle;
	}

	@Override
	public String getPromptBoxText() {
		return promptBoxText;
	}

	@Override
	public void setErrorBox(String title, String text) {
		errorBoxTitle = title;
		errorBoxText = text;
	}

	@Override
	public String getErrorBoxTitle() {
		return errorBoxTitle;
	}

	@Override
	public String getErrorBoxText() {
		return errorBoxText;
	}

	@Override
	public CellRegion getRegion() {
		return region;
	}
	
	@Override
	void setRegion(CellRegion region){
		Validations.argNotNull(region);
		this.region = region;
	}

	@Override
	public ValidationType getValidationType() {
		return validationType;
	}

	@Override
	public void setValidationType(ValidationType type) {
		Validations.argNotNull(type);
		validationType = type;
	}

	@Override
	public OperatorType getOperatorType() {
		return operatorType;
	}

	@Override
	public void setOperatorType(OperatorType type) {
		Validations.argNotNull(type);
		operatorType = type;
	}

	@Override
	public boolean isFormulaParsingError() {
		boolean r = false;
		if(value1Expr!=null){
			r |= value1Expr.hasError();
		}
		if(!r && value2Expr!=null){
			r |= value2Expr.hasError();
		}
		return r;
	}

	@Override
	public int getNumOfValue(){
		return getNumOfValue1();
	}
	@Override
	public Object getValue(int index) {
		return getValue1(index);
	}
	@Override
	public int getNumOfValue1(){
		evalFormula();
		return EvaluationUtil.sizeOf(evalValue1Result);
	}
	@Override
	public Object getValue1(int index) {
		evalFormula();
		if(index>=EvaluationUtil.sizeOf(evalValue1Result)){
			return null;
		}
		return EvaluationUtil.valueOf(evalValue1Result,index);
	}
	
	@Override
	public int getNumOfValue2(){
		evalFormula();
		return EvaluationUtil.sizeOf(evalValue2Result);
	}
	@Override
	public Object getValue2(int index) {
		evalFormula();
		if(index>=EvaluationUtil.sizeOf(evalValue2Result)){
			return null;
		}
		return EvaluationUtil.valueOf(evalValue2Result,index);
	}

	@Override
	public String getValueFormula() {
		return getValue1Formula();
	}
	
	@Override
	public String getValue1Formula() {
		return value1Expr==null?null:value1Expr.getFormulaString();
	}

	@Override
	public String getValue2Formula() {
		return value2Expr==null?null:value2Expr.getFormulaString();
	}

	private void clearFormulaDependency() {
		if(value1Expr!=null || value2Expr!=null){
			((AbstractBookSeriesAdv) sheet.getBook().getBookSeries())
					.getDependencyTable().clearDependents(getRef());
		}
	}
	
	private Ref getRef(){
		return new ObjectRefImpl(this,id);
	}
	
	@Override
	public void setFormula(String valueExpression) {
		setFormula(valueExpression,null);
	}
	@Override
	public void setFormula(String value1Expression, String value2Expression) {
		checkOrphan();
		evaluated = false;
		clearFormulaDependency();
		
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		
		Ref ref = getRef();
		if(value1Expression!=null){
			value1Expr = fe.parse(value1Expression, new FormulaParseContext(sheet,ref));
		}else{
			value1Expr = null;
		}
		if(value2Expression!=null){
			value2Expr = fe.parse(value2Expression, new FormulaParseContext(sheet,ref));
		}else{
			value2Expr = null;
		}
	}

	@Override
	public void clearFormulaResultCache() {
		evaluated = false;
		evalValue1Result = evalValue2Result = null;
	}
	
	/*package*/ void evalFormula(){
		if(!evaluated){
			FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
			if(value1Expr!=null){
				EvaluationResult result = fe.evaluate(value1Expr,new FormulaEvaluationContext(sheet));

				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalValue1Result = val;
				}else if(result.getType() == ResultType.ERROR){
					evalValue1Result = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_VALUE);
				}
				
			}
			if(value2Expr!=null){
				EvaluationResult result = fe.evaluate(value2Expr,new FormulaEvaluationContext(sheet));

				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					evalValue2Result = val;
				}else if(result.getType() == ResultType.ERROR){
					evalValue2Result = (val instanceof ErrorValue)?val:new ErrorValue(ErrorValue.INVALID_VALUE);
				}
				
			}
			evaluated = true;
		}
	}
	
	@Override
	public List<NCell> getReferToCellList(){
		if(value1Expr!=null && value1Expr.isRefersTo()){
			List<NCell> list = new LinkedList<NCell>();
			NBookSeries bookSeries = sheet.getBook().getBookSeries(); 
			
			String bookName = sheet.getBook().getBookName();//TODO zss 3.5 from expr
			String sheetName = value1Expr.getRefersToSheetName();
			CellRegion region = value1Expr.getRefersToCellRegion();
			
			NBook book = bookSeries.getBook(bookName);
			if(book==null){
				return list;
			}
			NSheet sheet = book.getSheetByName(sheetName);
			if(sheet==null){
				return list;
			}
			for(int i = region.getRow();i<=region.getLastRow();i++){
				for(int j=region.getColumn();j<=region.getLastColumn();j++){
					list.add(sheet.getCell(i, j));
				}
			}
			return list;
		}
		return Collections.EMPTY_LIST;
	}

	@Override
	public boolean hasReferToCellList() {
		return value1Expr!=null && value1Expr.isRefersTo();
	}

	@Override
	void copyFrom(AbstractDataValidationAdv src) {
		Validations.argInstance(src, DataValidationImpl.class);
		DataValidationImpl srcImpl = (DataValidationImpl)src;
		errorStyle = srcImpl.errorStyle;
		emptyCellAllowed = srcImpl.emptyCellAllowed;
		showDropDownArrow = srcImpl.showDropDownArrow;
		showPromptBox = srcImpl.showPromptBox;
		showErrorBox = srcImpl.showErrorBox;
		promptBoxTitle = srcImpl.promptBoxTitle;
		promptBoxText = srcImpl.promptBoxText;
		errorBoxTitle = srcImpl.errorBoxTitle;
		errorBoxText = srcImpl.errorBoxText;
		validationType = srcImpl.validationType;
		operatorType = srcImpl.operatorType;
		
		if(srcImpl.value1Expr!=null){
			setFormula(srcImpl.value1Expr==null?null:srcImpl.value1Expr.getFormulaString()
					, srcImpl.value2Expr==null?null:srcImpl.value2Expr.getFormulaString());
		}
	}

}
