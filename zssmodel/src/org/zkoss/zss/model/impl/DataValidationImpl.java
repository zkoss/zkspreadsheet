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

import java.util.ArrayList;
import java.util.Collections;
import java.util.HashSet;
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
import org.zkoss.zss.model.sys.dependency.ObjectRef.ObjectType;
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
	
	private AlertStyle _alertStyle = AlertStyle.STOP;//default stop
	private boolean _ignoreBlank = true;//default true
	private boolean _showInCellDropdown;
	private boolean _showInput;
	private boolean _showError;
	private String _inputTitle;
	private String _inputMessage;
	private String _errorTitle;
	private String _errorMessage;
	private Set<CellRegion> _regions;
	private ValidationType _validationType = ValidationType.ANY;
	private OperatorType _operatorType = OperatorType.BETWEEN;
	
	
	private FormulaExpression _formula1Expr;
	private FormulaExpression _formula2Expr;
	private Object _evalValue1Result;
	private Object _evalValue2Result;
	
	private boolean _evaluated = false;
	
	public DataValidationImpl(AbstractSheetAdv sheet,String id){
		this._sheet = sheet;
		this._id = id;
	}
	
	public boolean isEmpty() {
		return _validationType == ValidationType.ANY
				&& _inputTitle == null
				&& _inputMessage == null
				&& _errorTitle == null
				&& _errorMessage == null;
	}
	
	public DataValidationImpl(AbstractSheetAdv sheet, AbstractDataValidationAdv copyFrom) {
		this(sheet, (String) null);
		if (copyFrom != null) {
			this.copyFrom(copyFrom);
		}
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
	public AlertStyle getAlertStyle() {
		return _alertStyle;
	}

	@Override
	public void setAlertStyle(AlertStyle alertStyle) {
		Validations.argNotNull(alertStyle);
		this._alertStyle = alertStyle;
	}

	@Override
	public void setIgnoreBlank(boolean allowed) {
		this._ignoreBlank = allowed;
	}

	@Override
	public boolean isIgnoreBlank() {
		return _ignoreBlank;
	}

	@Override
	public void setInCellDropdown(boolean show) {
		_showInCellDropdown = show;
	}

	@Override
	public boolean isInCellDropdown() {
		return _showInCellDropdown;
	}

	@Override
	public void setShowInput(boolean show) {
		_showInput = show;
	}

	@Override
	public boolean isShowInput() {
		return _showInput;
	}

	@Override
	public void setShowError(boolean show) {
		_showError = show;
	}

	@Override
	public boolean isShowError() {
		return _showError;
	}

	@Override
	public void setInputTitle(String title) {
		_inputTitle = title;
	}
	@Override
	public void setInputMessage(String message) {
		_inputMessage = message;
	}

	@Override
	public String getInputTitle() {
		return _inputTitle;
	}

	@Override
	public String getInputMessage() {
		return _inputMessage;
	}

	@Override
	public void setErrorTitle(String title) {
		_errorTitle = title;
	}
	
	@Override
	public void setErrorMessage(String text) {
		_errorMessage = text;
	}

	@Override
	public String getErrorTitle() {
		return _errorTitle;
	}

	@Override
	public String getErrorMessage() {
		return _errorMessage;
	}

	@Override
	public Set<CellRegion> getRegions() {
		return _regions;
	}
	
	@Override
	public void addRegion(CellRegion region) { // ZSS-648
		Validations.argNotNull(region);
		if (this._regions == null) {
			this._regions = new HashSet<CellRegion>(2);
		}
		for (CellRegion regn : this._regions) {
			if (regn.contains(region)) return; // already in this DataValidation, let go
		}
		
		this._regions.add(region);
		
		// ZSS-648
		// Add new ObjectRef into DependencyTable so we can extend/shrink/move
		Ref dependent = getRef();
		SBook book = _sheet.getBook();
		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		// prepare a dummy CellRef to enforce DataValidation reference dependency
		dt.add(dependent, newDummyRef(region));
		
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	@Override
	public void removeRegion(CellRegion region) { // ZSS-694
		Validations.argNotNull(region);
		if (this._regions == null || this._regions.isEmpty()) return;
		
		List<CellRegion> newRegions = new ArrayList<CellRegion>();
		List<CellRegion> delRegions = new ArrayList<CellRegion>();
		for (CellRegion regn : this._regions) {
			if (!regn.overlaps(region)) continue;
			newRegions.addAll(regn.diff(region));
			delRegions.add(regn);
		}
		
		// no overlapping at all
		if (newRegions.isEmpty() && delRegions.isEmpty()) {
			return;
		}

		Ref dependent = getRef();
		SBook book = _sheet.getBook();
		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		final Set<Ref> precedents = ((DependencyTableAdv)dt).getDirectPrecedents(dependent);
		dt.clearDependents(dependent);
		
		for (CellRegion regn : delRegions) {
			this._regions.remove(regn);
			precedents.remove(newDummyRef(regn));
		}
		
		// restore dependent precedents relation
		if (precedents != null) {
			for (Ref precedent: precedents) {
				dt.add(dependent, precedent);
			}
		}
		
		// add new split regions
		for (CellRegion regn : newRegions) {
			this._regions.add(regn);
			dt.add(dependent, newDummyRef(regn));
		}
		
		if (this._regions.isEmpty()) {
			this._regions = null;
		}
		
		ModelUpdateUtil.addRefUpdate(dependent);
	}
	
	@Override
	public void setRegions(Set<CellRegion> regions) {
		_regions = new HashSet<CellRegion>(regions.size() * 4 / 3 + 1);
		for (CellRegion rgn : regions) {
			addRegion(rgn);
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
		if(_formula1Expr!=null){
			r |= _formula1Expr.hasError();
		}
		if(!r && _formula2Expr!=null){
			r |= _formula2Expr.hasError();
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
	public String getFormula1() {
		return _formula1Expr==null?null:_formula1Expr.getFormulaString();
	}

	@Override
	public String getFormula2() {
		return _formula2Expr==null?null:_formula2Expr.getFormulaString();
	}

	private void clearFormulaDependency(boolean all) { // ZSS-648
		if(_formula1Expr!=null || _formula2Expr!=null){
			Ref dependent = getRef();
			DependencyTable dt = 
			((AbstractBookSeriesAdv) _sheet.getBook().getBookSeries())
					.getDependencyTable();
			
			dt.clearDependents(dependent);
			
			// ZSS-648
			// must keep the region reference itself in DependencyTable; so add it back
			if (!all && this._regions != null) {
				for (CellRegion regn : this._regions) {
					dt.add(dependent, newDummyRef(regn));
				}
			}
		}
	}
	
	private Ref getRef(){
		return new ObjectRefImpl(this,_id);
	}
	
	// ZSS-648
	private Ref getRef(String sheetName) {
		return new ObjectRefImpl(_sheet.getBook().getBookName(), sheetName, _id, ObjectType.DATA_VALIDATION);
	}
	
	// ZSS-648
	private Ref newDummyRef(CellRegion regn) {
		return new RefImpl(_sheet.getBook().getBookName(), _sheet.getSheetName(), 
				regn.row, regn.column, regn.lastRow, regn.lastColumn);
	}
	
	// ZSS-648
	private Ref newDummyRef(String sheetName, CellRegion regn) {
		return new RefImpl(_sheet.getBook().getBookName(), sheetName, 
				regn.row, regn.column, regn.lastRow, regn.lastColumn);
	}
	
	@Override
	public void setFormula1(String formula1) {
		setFormulas(formula1, _formula2Expr == null ? null : _formula2Expr.getFormulaString());
	}
	
	@Override
	public void setFormula2(String formula2) {
		setFormulas(_formula1Expr == null ? null : _formula1Expr.getFormulaString(), formula2);
	}
	
	@Override
	public void setFormulas(String formula1, String formula2) {
		checkOrphan();
		_evaluated = false;
		clearFormulaDependency(false); // will clear formula
		clearFormulaResultCache();
		
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		
		Ref ref = getRef();
		if(formula1!=null){
			_formula1Expr = fe.parse(formula1, new FormulaParseContext(_sheet,ref));
		}else{
			_formula1Expr = null;
		}
		
		if(formula2!=null){
			_formula2Expr = fe.parse(formula2, new FormulaParseContext(_sheet,ref));
		}else{
			_formula2Expr = null;
		}
	}

	@Override
	public void clearFormulaResultCache() {
		_evaluated = false;
		_evalValue1Result = _evalValue2Result = null;
	}
	
	/*package*/ void evalFormula(){
		//20140731, henrichen: when share the same book, many users might 
		//populate DataValidationImpl simultaneously; must synchronize it.
		if(_evaluated) return;
		synchronized (this) {
			if(!_evaluated){
				Ref ref = getRef();
				FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
				if(_formula1Expr!=null){
					EvaluationResult result = fe.evaluate(_formula1Expr,new FormulaEvaluationContext(_sheet,ref));
	
					Object val = result.getValue();
					if(result.getType() == ResultType.SUCCESS){
						_evalValue1Result = val;
					}else if(result.getType() == ResultType.ERROR){
						_evalValue1Result = (val instanceof ErrorValue)?val: ErrorValue.valueOf(ErrorValue.INVALID_VALUE);
					}
					
				}
				if(_formula2Expr!=null){
					EvaluationResult result = fe.evaluate(_formula2Expr,new FormulaEvaluationContext(_sheet,ref));
	
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
	}
	
	@Override
	public List<SCell> getReferToCellList(){
		if(_formula1Expr!=null && _formula1Expr.isAreaRefs()){
			List<SCell> list = new LinkedList<SCell>();
			SBookSeries bookSeries = _sheet.getBook().getBookSeries(); 
			
			Ref areaRef = _formula1Expr.getAreaRefs()[0];
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
		return Collections.emptyList();
	}

	@Override
	public boolean hasReferToCellList() {
		return _formula1Expr!=null && _formula1Expr.isAreaRefs();
	}

	@Override
	void copyFrom(AbstractDataValidationAdv src) {
		Validations.argInstance(src, DataValidationImpl.class);
		DataValidationImpl srcImpl = (DataValidationImpl)src;
		_alertStyle = srcImpl._alertStyle;
		_ignoreBlank = srcImpl._ignoreBlank;
		_showInCellDropdown = srcImpl._showInCellDropdown;
		_showInput = srcImpl._showInput;
		_showError = srcImpl._showError;
		_inputTitle = srcImpl._inputTitle;
		_inputMessage = srcImpl._inputMessage;
		_errorTitle = srcImpl._errorTitle;
		_errorMessage = srcImpl._errorMessage;
		_validationType = srcImpl._validationType;
		_operatorType = srcImpl._operatorType;
		
		if(srcImpl._formula1Expr!=null){
			setFormulas(srcImpl._formula1Expr==null?null:srcImpl._formula1Expr.getFormulaString(), 
						srcImpl._formula2Expr==null?null:srcImpl._formula2Expr.getFormulaString());
		}
	}
	
	@Override
	void renameSheet(String oldName, String newName) { //ZSS-648
		Validations.argNotNull(oldName);
		Validations.argNotNull(newName);
		if (oldName.equals(newName)) return; // nothing change, let go
		
		// remove old ObjectRef
		Ref dependent = getRef(oldName);
		SBook book = _sheet.getBook();
		final DependencyTable dt = 
				((AbstractBookSeriesAdv) book.getBookSeries()).getDependencyTable();
		final Set<Ref> precedents = ((DependencyTableAdv)dt).getDirectPrecedents(dependent);
		if (precedents != null && this._regions != null) {
			for (CellRegion regn : this._regions) {
				precedents.remove(newDummyRef(oldName, regn));
			}
		}
		dt.clearDependents(dependent);
		
		// Add new ObjectRef into DependencyTable so we can extend/shrink/move
		dependent = getRef(newName);  // new dependent because region have changed
		
		// ZSS-648
		// prepare new dummy CellRef to enforce DataValidation reference dependency
		if (this._regions != null) {
			for (CellRegion regn : this._regions) {
				dt.add(dependent, newDummyRef(newName, regn));
			}
		}
		
		// restore dependent precedents relation
		if (precedents != null) {
			for (Ref precedent: precedents) {
				dt.add(dependent, precedent);
			}
		}
	}

	//ZSS-688
	//@since 3.6.0
	/*package*/ DataValidationImpl cloneDataValidationImpl(AbstractSheetAdv sheet) {
		DataValidationImpl tgt = new DataValidationImpl(sheet, this._id); 

		tgt._alertStyle = this._alertStyle;
		tgt._ignoreBlank = this._ignoreBlank;
		tgt._showInCellDropdown = this._showInCellDropdown;
		tgt._showInput = this._showInput;
		tgt._showError = this._showError;
		tgt._inputTitle = this._inputTitle;
		tgt._inputMessage = this._inputMessage;
		tgt._errorTitle = this._errorTitle;
		tgt._errorMessage = this._errorMessage;
		tgt._validationType = this._validationType;
		tgt._operatorType = this._operatorType;

		if (this._regions != null) {
			tgt._regions = new HashSet<CellRegion>(this._regions.size() * 4 / 3);
			for (CellRegion rgn : this._regions) {
				tgt._regions.add(new CellRegion(rgn.row, rgn.column, rgn.lastRow, rgn.lastColumn));
			}
		}
		
		final String f1 = getFormula1();
		if (f1 != null) {
			final String f2 = getFormula2();
			setFormulas(f1, f2);
		}

		// Do NOT clone _evalValue1Result, _evalValue2Result, and _evaluated
		//private Object _evalValue1Result;
		//private Object _evalValue2Result;
		//private boolean _evaluated;
		
		return tgt;
	}
	
	//ZSS-747
	/**
	 * 
	 * @param fe1
	 * @param fe2
	 * @since 3.6.0
	 */
	public void setFormulas(FormulaExpression fe1, FormulaExpression fe2) {
		checkOrphan();
		_evaluated = false;
		clearFormulaDependency(false); // will clear formula
		clearFormulaResultCache();
		
		_formula1Expr = fe1;
		_formula2Expr = fe2;

		// update dependency table
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		
		Ref ref = getRef();
		FormulaParseContext context = new FormulaParseContext(_sheet,ref);
		
		if(fe1 != null) {
			fe.updateDependencyTable(fe1, context);
		}
		if(fe2 != null) {
			fe.updateDependencyTable(fe2, context);
		}
	}

	//ZSS-747
	/**
	 * 
	 * @return
	 * @since 3.6.0
	 */
	public FormulaExpression getFormulaExpression1() {
		return _formula1Expr;
	}
	//ZSS-747
	/**
	 * 
	 * @return
	 * @since 3.6.0
	 */
	public FormulaExpression getFormulaExpression2() {
		return _formula2Expr;
	}
	//ZSS-747
	/**
	 * 
	 * @param formula
	 * @since 3.6.0
	 */
	public void setFormula1(FormulaExpression formula1) {
		setFormulas(formula1, _formula2Expr);
	}
	//ZSS-747
	/**
	 * 
	 * @param formula
	 * @since 3.6.0
	 */
	public void setFormula2(FormulaExpression formula2) {
		setFormulas(_formula1Expr, formula2);
	}
}
