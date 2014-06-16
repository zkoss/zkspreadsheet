/*

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.model.impl;

import java.util.Collections;
import java.util.LinkedList;
import java.util.List;
import java.util.Set;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.ErrorValue;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SBookSeries;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.DependencyTable;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.formula.EvaluationResult;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
import org.zkoss.zss.model.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.model.util.Validations;
import org.zkoss.zss.model.impl.sys.DependencyTableAdv;
/**
 * 
 * @author Dennis
 * @since 3.5.0
 */
public class DataValidationImpl extends AbstractDataValidationAdv {

	private static final long serialVersionUID = 1L;
	private AbstractSheetAdv _sheet;
	final private String _id;
	
	private ErrorStyle _errorStyle = ErrorStyle.STOP;//default stop
	private boolean _emptyCellAllowed = true;//default true
	private boolean _showDropDownArrow;
	private boolean _showPromptBox;
	private boolean _showErrorBox;
	private String _promptBoxTitle;
	private String _promptBoxText;
	private String _errorBoxTitle;
	private String _errorBoxText;
	private CellRegion _region;
	private ValidationType _validationType = ValidationType.ANY;
	private OperatorType _operatorType = OperatorType.BETWEEN;
	
	
	private FormulaExpression _value1Expr;
	private FormulaExpression _value2Expr;
	private Object _evalValue1Result;
	private Object _evalValue2Result;
	
	private boolean _evaluated = false;
	
	public DataValidationImpl(AbstractSheetAdv sheet,String id){
		this._sheet = sheet;
		this._id = id;
	}
	
	public String getId(){
		return _id;
	}
	
	public SSheet getSheet(){
		return _sheet;
	}
	
	@Override
	public void checkOrphan() {
		if (_sheet == null) {
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	public void destroy() {
		checkOrphan();
		clearFormulaDependency(true);
		clearFormulaResultCache();
		_sheet = null;
	}
	
	@Override
	public ErrorStyle getErrorStyle() {
		return _errorStyle;
	}

	@Override
	public void setErrorStyle(ErrorStyle errorStyle) {
		Validations.argNotNull(errorStyle);
		this._errorStyle = errorStyle;
	}

	@Override
	public void setEmptyCellAllowed(boolean allowed) {
		this._emptyCellAllowed = allowed;
	}

	@Override
	public boolean isEmptyCellAllowed() {
		return _emptyCellAllowed;
	}

	@Override
	public void setShowDropDownArrow(boolean show) {
		_showDropDownArrow = show;
	}

	@Override
	public boolean isShowDropDownArrow() {
		return _showDropDownArrow;
	}

	@Override
	public void setShowPromptBox(boolean show) {
		_showPromptBox = show;
	}

	@Override
	public boolean isShowPromptBox() {
		return _showPromptBox;
	}

	@Override
	public void setShowErrorBox(boolean show) {
		_showErrorBox = show;
	}

	@Override
	public boolean isShowErrorBox() {
		return _showErrorBox;
	}

	@Override
	public void setPromptBox(String title, String text) {
		_promptBoxTitle = title;
		_promptBoxText = text;
	}

	@Override
	public String getPromptBoxTitle() {
		return _promptBoxTitle;
	}

	@Override
	public String getPromptBoxText() {
		return _promptBoxText;
	}

	@Override
	public void setErrorBox(String title, String text) {
		_errorBoxTitle = title;
		_errorBoxText = text;
	}

	@Override
	public String getErrorBoxTitle() {
		return _errorBoxTitle;
	}

	@Override
	public String getErrorBoxText() {
		return _errorBoxText;
	}

	@Override
	public CellRegion getRegion() {
		return _region;
	}
	
	@Override
	void setRegion(CellRegion region) { // ZSS-648
		Validations.argNotNull(region);
		if (region.equals(this._region)) return; // nothing change, let go
		
		// ZSS-648
		// remove old ObjectRef
		Ref dependent = getRef();
		Ref dummy = getDummyRef();
		SBook book = _sheet.getBook();
		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		final Set<Ref> precedents = ((DependencyTableAdv)dt).getDirectPrecedents(dependent);
		if (precedents != null) precedents.remove(dummy);
		dt.clearDependents(dependent);
		
		// Add new ObjectRef into DependencyTable so we can extend/shrink/move
		this._region = region;
		dependent = getRef();  // new dependent because region have changed
		
		// ZSS-648
		// prepare a dummy CellRef to enforce DataValidation reference dependency
		dummy = getDummyRef();
		if (dummy != null) dt.add(dependent, dummy);
		
		// restore dependent precedents relation
		if (precedents != null) {
			for (Ref precedent: precedents) {
				dt.add(dependent, precedent);
			}
		}
	}

	@Override
	public ValidationType getValidationType() {
		return _validationType;
	}

	@Override
	public void setValidationType(ValidationType type) {
		Validations.argNotNull(type);
		_validationType = type;
	}

	@Override
	public OperatorType getOperatorType() {
		return _operatorType;
	}

	@Override
	public void setOperatorType(OperatorType type) {
		Validations.argNotNull(type);
		_operatorType = type;
	}

	@Override
	public boolean isFormulaParsingError() {
		boolean r = false;
		if(_value1Expr!=null){
			r |= _value1Expr.hasError();
		}
		if(!r && _value2Expr!=null){
			r |= _value2Expr.hasError();
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
		return EvaluationUtil.sizeOf(_evalValue1Result);
	}
	@Override
	public Object getValue1(int index) {
		evalFormula();
		if(index>=EvaluationUtil.sizeOf(_evalValue1Result)){
			return null;
		}
		return EvaluationUtil.valueOf(_evalValue1Result,index);
	}
	
	@Override
	public int getNumOfValue2(){
		evalFormula();
		return EvaluationUtil.sizeOf(_evalValue2Result);
	}
	@Override
	public Object getValue2(int index) {
		evalFormula();
		if(index>=EvaluationUtil.sizeOf(_evalValue2Result)){
			return null;
		}
		return EvaluationUtil.valueOf(_evalValue2Result,index);
	}

	@Override
	public String getValueFormula() {
		return getValue1Formula();
	}
	
	@Override
	public String getValue1Formula() {
		return _value1Expr==null?null:_value1Expr.getFormulaString();
	}

	@Override
	public String getValue2Formula() {
		return _value2Expr==null?null:_value2Expr.getFormulaString();
	}

	private void clearFormulaDependency(boolean all) { // ZSS-648
		if(_value1Expr!=null || _value2Expr!=null){
			Ref dependent = getRef();
			DependencyTable dt = 
			((AbstractBookSeriesAdv) _sheet.getBook().getBookSeries())
					.getDependencyTable();
			
			dt.clearDependents(dependent);
			
			// ZSS-648
			// must keep the region reference itself in DependencyTable; so add it back
			if (!all) {
				final Ref dummy = getDummyRef();
				if (dummy != null) dt.add(dependent, dummy);
			}
		}
	}
	
	private Ref getRef(){
		return new ObjectRefImpl(this,_id);
	}
	
	// ZSS-648
	private Ref getRef(String sheetName) {
		return new ObjectRefImpl(_sheet.getBook().getBookName(), sheetName, _id);
	}
	
	// ZSS-648
	private Ref getDummyRef() {
		return _region == null ? null :
			new RefImpl(_sheet.getBook().getBookName(), _sheet.getSheetName(), 
				_region.row, _region.column, _region.lastRow, _region.lastColumn);
	}
	
	// ZSS-648
	private Ref getDummyRef(String sheetName) {
		return _region == null ? null :
			new RefImpl(_sheet.getBook().getBookName(), sheetName, 
				_region.row, _region.column, _region.lastRow, _region.lastColumn);
	}
	
	@Override
	public void setFormula(String valueExpression) {
		setFormula(valueExpression,null);
	}
	@Override
	public void setFormula(String value1Expression, String value2Expression) {
		checkOrphan();
		_evaluated = false;
		clearFormulaDependency(false);
		clearFormulaResultCache();
		
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		
		Ref ref = getRef();
		if(value1Expression!=null){
			_value1Expr = fe.parse(value1Expression, new FormulaParseContext(_sheet,ref));
		}else{
			_value1Expr = null;
		}
		if(value2Expression!=null){
			_value2Expr = fe.parse(value2Expression, new FormulaParseContext(_sheet,ref));
		}else{
			_value2Expr = null;
		}
	}

	@Override
	public void clearFormulaResultCache() {
		_evaluated = false;
		_evalValue1Result = _evalValue2Result = null;
	}
	
	/*package*/ void evalFormula(){
		if(!_evaluated){
			Ref ref = getRef();
			FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
			if(_value1Expr!=null){
				EvaluationResult result = fe.evaluate(_value1Expr,new FormulaEvaluationContext(_sheet,ref));

				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					_evalValue1Result = val;
				}else if(result.getType() == ResultType.ERROR){
					_evalValue1Result = (val instanceof ErrorValue)?val: ErrorValue.valueOf(ErrorValue.INVALID_VALUE);
				}
				
			}
			if(_value2Expr!=null){
				EvaluationResult result = fe.evaluate(_value2Expr,new FormulaEvaluationContext(_sheet,ref));

				Object val = result.getValue();
				if(result.getType() == ResultType.SUCCESS){
					_evalValue2Result = val;
				}else if(result.getType() == ResultType.ERROR){
					_evalValue2Result = (val instanceof ErrorValue)?val: ErrorValue.valueOf(ErrorValue.INVALID_VALUE);
				}
				
			}
			_evaluated = true;
		}
	}
	
	@Override
	public List<SCell> getReferToCellList(){
		if(_value1Expr!=null && _value1Expr.isAreaRefs()){
			List<SCell> list = new LinkedList<SCell>();
			SBookSeries bookSeries = _sheet.getBook().getBookSeries(); 
			
			Ref areaRef = _value1Expr.getAreaRefs()[0];
			String bookName =  areaRef.getBookName();
			String sheetName = areaRef.getSheetName();
			CellRegion region = new CellRegion(areaRef.getRow(),areaRef.getColumn(),areaRef.getLastRow(),areaRef.getLastColumn());
			
			SBook book = bookSeries.getBook(bookName);
			if(book==null){
				return list;
			}
			SSheet sheet = book.getSheetByName(sheetName);
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
		return _value1Expr!=null && _value1Expr.isAreaRefs();
	}

	@Override
	void copyFrom(AbstractDataValidationAdv src) {
		Validations.argInstance(src, DataValidationImpl.class);
		DataValidationImpl srcImpl = (DataValidationImpl)src;
		_errorStyle = srcImpl._errorStyle;
		_emptyCellAllowed = srcImpl._emptyCellAllowed;
		_showDropDownArrow = srcImpl._showDropDownArrow;
		_showPromptBox = srcImpl._showPromptBox;
		_showErrorBox = srcImpl._showErrorBox;
		_promptBoxTitle = srcImpl._promptBoxTitle;
		_promptBoxText = srcImpl._promptBoxText;
		_errorBoxTitle = srcImpl._errorBoxTitle;
		_errorBoxText = srcImpl._errorBoxText;
		_validationType = srcImpl._validationType;
		_operatorType = srcImpl._operatorType;
		
		if(srcImpl._value1Expr!=null){
			setFormula(srcImpl._value1Expr==null?null:srcImpl._value1Expr.getFormulaString()
					, srcImpl._value2Expr==null?null:srcImpl._value2Expr.getFormulaString());
		}
	}
	
	@Override
	void renameSheet(String oldName, String newName) { //ZSS-648
		Validations.argNotNull(oldName);
		Validations.argNotNull(newName);
		if (oldName.equals(newName)) return; // nothing change, let go
		
		// remove old ObjectRef
		Ref dependent = getRef(oldName);
		Ref dummy = getDummyRef(oldName);
		SBook book = _sheet.getBook();
		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		final Set<Ref> precedents = ((DependencyTableAdv)dt).getDirectPrecedents(dependent);
		if (precedents != null) precedents.remove(dummy);
		dt.clearDependents(dependent);
		
		// Add new ObjectRef into DependencyTable so we can extend/shrink/move
		dependent = getRef(newName);  // new dependent because region have changed
		
		// ZSS-648
		// prepare a dummy CellRef to enforce DataValidation reference dependency
		dummy = getDummyRef(newName);
		if (dummy != null) dt.add(dependent, dummy);
		
		// restore dependent precedents relation
		if (precedents != null) {
			for (Ref precedent: precedents) {
				dt.add(dependent, precedent);
			}
		}
	}
}
