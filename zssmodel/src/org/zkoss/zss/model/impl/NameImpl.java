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
import org.zkoss.zss.model.sys.formula.FormulaExpression;
import org.zkoss.zss.model.sys.formula.FormulaParseContext;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NameImpl extends AbstractNameAdv {

	private final String _id;
	private AbstractBookAdv _book;
	private String _name;
	
	private String _applyToSheetName;
	
	private String _refersToExpr;
	
	private CellRegion _refersToCellRegion;
	private String _refersTopSheetName;
	
	private boolean _isParsingError;
	
	public NameImpl(AbstractBookAdv book, String id) {
		this._book = book;
		this._id = id;
	}

	@Override
	public String getName() {
		return _name;
	}

	@Override
	public String getRefersToSheetName() {
		return _refersTopSheetName;
	}

	@Override
	public CellRegion getRefersToCellRegion() {
		return _refersToCellRegion;
	}

	@Override
	public String getRefersToFormula() {
		return _refersToExpr;
	}

	@Override
	public void destroy() {
		checkOrphan();
		clearFormulaDependency();
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
	void setName(String newname) {
		_name = newname;
	}

	@Override
	public String getId() {
		return _id;
	}

	@Override
	public void setRefersToFormula(String refersToExpr) {
		checkOrphan();
		
		clearFormulaDependency();
		this._refersToExpr = refersToExpr;
		_refersTopSheetName = null;
		_refersToCellRegion = null;
		//TODO support function as Excel (POI)
		
		//use formula engine to keep dependency info
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		FormulaExpression expr = fe.parse(refersToExpr, new FormulaParseContext(_book.getSheet(0),new NameRefImpl(this)));
		if(expr.hasError()){
			_isParsingError = true;
		}else if(expr.isRefersTo()){
			_refersTopSheetName = expr.getRefersToSheetName();
			_refersToCellRegion = expr.getRefersToCellRegion();
		}
		
	}
	
	@Override
	public boolean isFormulaParsingError() {
		return _isParsingError;
	}

	private void clearFormulaDependency() {
		if(_refersToExpr!=null){
			Ref ref = new NameRefImpl(this);
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

	@Override
	void setApplyToSheetName(String sheetName) {
		_applyToSheetName = sheetName;
	}
}
