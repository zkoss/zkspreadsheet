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
package org.zkoss.zss.model.impl;

import org.zkoss.zss.model.CellRegion;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.dependency.Ref;
import org.zkoss.zss.model.sys.formula.FormulaEngine;
import org.zkoss.zss.model.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NameImpl extends AbstractNameAdv {
	private static final long serialVersionUID = 1L;
	private final String _id;
	private AbstractBookAdv _book;
	private String _name;
	
	private String _applyToSheetName;
	
	private String _refersToExpr;
	
	private FormulaExpression _refersToExprFormula;
	
	private CellRegion _refersToCellRegion;
	private String _refersToSheetName;
	
	public NameImpl(AbstractBookAdv book, String id, String name, String applyToSheetName) {
		this._book = book;
		this._id = id;
		this._name = name;
		this._applyToSheetName = applyToSheetName;
	}

	@Override
	public String getName() {
		return _name;
	}

	@Override
	public String getRefersToSheetName() {
		return _refersToSheetName;
	}

	@Override
	public CellRegion getRefersToCellRegion() {
		return _refersToCellRegion;
	}

	@Override
	public String getRefersToFormula() {
		return _refersToExprFormula==null?null:_refersToExprFormula.getFormulaString();
	}

	@Override
	public void destroy() {
		checkOrphan();
		clearFormulaDependency(_name); // ZSS-661
		clearFormulaResultCache();
		_book = null;
	}

	@Override
	public void checkOrphan() {
		if(_book==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	void setName(String newname, String applyToSheetName) {
		checkOrphan();
		_applyToSheetName = applyToSheetName;
		parseAndEvalFormula(_name, newname); // ZSS-661
		_name = newname;
	}

	@Override
	public String getId() {
		return _id;
	}

	@Override
	public void setRefersToFormula(String refersToExpr) {
		checkOrphan();
		_refersToExpr = refersToExpr;
		parseAndEvalFormula(_name, _name); // nameName not changed
	}
	
	/*package*/ void parseAndEvalFormula(String oldName, String newName){ //ZSS-661
		clearFormulaDependency(oldName);

		_refersToSheetName = null;
		_refersToCellRegion = null;
		
		if(_refersToExpr!=null){
			//use formula engine to keep dependency info
			FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
			Ref ref = getRef(newName); // ZSS-661
			_refersToExprFormula = fe.parse(_refersToExpr, new FormulaParseContext(_book.getSheet(0),ref));
			if(!_refersToExprFormula.hasError() && _refersToExprFormula.isAreaRefs()){
				//ZSS-655, should eval each name to force it be notified in dependency (for evaluated dependent)
				//we don't care result here
				fe.evaluate(_refersToExprFormula,new FormulaEvaluationContext(_book.getSheet(0), ref));
				
				//TODO, should care all the refs
				Ref[] refs = _refersToExprFormula.getAreaRefs();
				_refersToSheetName = refs[0].getSheetName();
				_refersToCellRegion = new CellRegion(refs[0].getRow(),refs[0].getColumn(),refs[0].getLastRow(),refs[0].getLastColumn());				
				
				//ZSS-294, should clear dependents that have referenced to this new created Name
				ModelUpdateUtil.handlePrecedentUpdate(_book.getBookSeries(), ref);
			}
		}
	}	
	
	@Override
	public boolean isFormulaParsingError() {
		return _refersToExprFormula==null?false:_refersToExprFormula.hasError();
	}

	private void clearFormulaDependency(String nameName) {
		if(_refersToExprFormula!=null){
			Ref ref = getRef(nameName); // ZSS-661
			((AbstractBookSeriesAdv)_book.getBookSeries()).getDependencyTable().clearDependents(ref);
		}
	}

	@Override
	public AbstractBookAdv getBook() {
		checkOrphan();
		return _book;
	}

	@Override
	public void clearFormulaResultCache() {
		//so far, no result cache here, do nothing
	}

	@Override
	public String getApplyToSheetName() {
		return _applyToSheetName;
	}
	
	private Ref getRef(String nameName) { // ZSS-661
		return new NameRefImpl(this.getBook().getBookName(),this.getApplyToSheetName(),nameName);
	}
}
