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

import org.zkoss.zss.ngmodel.CellRegion;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NameImpl extends NameAdv {

	private final String id;
	private BookAdv book;
	private String name;
	
	private String applyToSheetName;
	
	private String refersToExpr;
	
	private CellRegion refersToCellRegion;
	private String refersTopSheetName;
	
	private boolean isParsingError;
	
	public NameImpl(BookAdv book, String id) {
		this.book = book;
		this.id = id;
	}

	@Override
	public String getName() {
		return name;
	}

	@Override
	public String getRefersToSheetName() {
		return refersTopSheetName;
	}

	@Override
	public CellRegion getRefersToCellRegion() {
		return refersToCellRegion;
	}

	@Override
	public String getRefersToFormula() {
		return refersToExpr;
	}

	@Override
	public void destroy() {
		checkOrphan();
		clearFormulaDependency();
		clearFormulaResultCache();
		book = null;
	}

	@Override
	public void checkOrphan() {
		if(book==null){
			throw new IllegalStateException("doesn't connect to parent");
		}
	}

	@Override
	void setName(String newname) {
		name = newname;
	}

	@Override
	public String getId() {
		return id;
	}

	@Override
	public void setRefersToFormula(String refersToExpr) {
		checkOrphan();
		
		clearFormulaDependency();
		this.refersToExpr = refersToExpr;
		refersTopSheetName = null;
		refersToCellRegion = null;
		//TODO support function as Excel (POI)
		
		//use formula engine to keep dependency info
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		FormulaExpression expr = fe.parse(refersToExpr, new FormulaParseContext(book.getSheet(0),new NameRefImpl(this)));
		if(expr.hasError()){
			isParsingError = true;
		}else if(expr.isRefersTo()){
			refersTopSheetName = expr.getRefersToSheetName();
			refersToCellRegion = expr.getRefersToCellRegion();
		}
		
	}
	
	@Override
	public boolean isFormulaParsingError() {
		return isParsingError;
	}

	private void clearFormulaDependency() {
		if(refersToExpr!=null){
			Ref ref = new NameRefImpl(this);
			((BookSeriesAdv)book.getBookSeries()).getDependencyTable().clearDependents(ref);
		}
	}

	@Override
	public BookAdv getBook() {
		checkOrphan();
		return book;
	}

	@Override
	public void clearFormulaResultCache() {
		//so far, no result cache here, do nothing
	}

	@Override
	public String getApplyToSheetName() {
		return applyToSheetName;
	}

	@Override
	void setApplyToSheetName(String sheetName) {
		applyToSheetName = sheetName;
	}
}
