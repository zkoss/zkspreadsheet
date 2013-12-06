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
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult;
import org.zkoss.zss.ngmodel.sys.formula.EvaluationResult.ResultType;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEvaluationContext;
import org.zkoss.zss.ngmodel.sys.formula.FormulaExpression;
import org.zkoss.zss.ngmodel.sys.formula.FormulaParseContext;
import org.zkoss.zss.ngmodel.util.AreaReference;
import org.zkoss.zss.ngmodel.util.CellReference;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class NameImpl extends NameAdv {

	private final String id;
	private BookAdv book;
	private String name;
	
	private String refersToExpr;
	
	private CellRegion refersTo;
	private String sheetName;
	
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
	public String getSheetName() {
		return sheetName;
	}

	@Override
	public CellRegion getRefersTo() {
		return refersTo;
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
		sheetName = null;
		refersTo = null;
		//TODO support function as Excel (POI)
		
		//use formula engine to keep dependency info
		FormulaEngine fe = EngineFactory.getInstance().createFormulaEngine();
		FormulaExpression expr = fe.parse(refersToExpr, new FormulaParseContext(book,new ObjectRefImpl(this,id)));
		if(expr.hasError()){
			isParsingError = true;
		}else{
			EvaluationResult result = fe.evaluate(expr,new FormulaEvaluationContext(book));
			if(result.getType()==ResultType.ERROR){
				isParsingError = true;
			}
		}
		
		
		if(!isParsingError){
			AreaReference ref = new AreaReference(refersToExpr);
			CellReference cell1 = ref.getFirstCell();
			CellReference cell2 = ref.getLastCell();
			sheetName = cell1.getSheetName();
			refersTo = new CellRegion(cell1.getRow(),cell1.getCol(),cell2.getRow(),cell2.getCol());
		}
		
	}
	
	@Override
	public boolean isFormulaParsingError() {
		return isParsingError;
	}

	private void clearFormulaDependency() {
		if(refersToExpr!=null){
			Ref ref = new ObjectRefImpl(this,id);
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
}
