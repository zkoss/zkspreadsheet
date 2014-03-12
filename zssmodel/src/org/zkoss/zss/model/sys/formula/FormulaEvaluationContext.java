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
package org.zkoss.zss.model.sys.formula;

import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.sys.AbstractContext;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class FormulaEvaluationContext extends AbstractContext {

	private final SBook _book;
	private final SSheet _sheet;
	private final SCell _cell;
	private FunctionMapper _functionMapper;
	private VariableResolver _vairableResolver;

	public FormulaEvaluationContext(SCell cell) {
		this(cell.getSheet().getBook(), cell.getSheet(), cell);
	}

	public FormulaEvaluationContext(SSheet sheet) {
		this(sheet.getBook(), sheet, null);
	}

	public FormulaEvaluationContext(SBook book) {
		this(book, null, null);
	}

	private FormulaEvaluationContext(SBook book, SSheet sheet, SCell cell) {
		this._book = book;
		this._sheet = sheet;
		this._cell = cell;
		EvaluationContributor contributor = book instanceof EvaluationContributorContainer? 
				((EvaluationContributorContainer)book).getEvaluationContributor():null;
		if(contributor!=null){
			_functionMapper = contributor.getFunctionMaper(book);
			_vairableResolver = contributor.getVariableResolver(book);
		}
	}

	public SBook getBook() {
		return _book;
	}

	public SSheet getSheet() {
		return _sheet;
	}

	public SCell getCell() {
		return _cell;
	}
	
	public FunctionMapper getFunctionMapper() {
		return _functionMapper;
	}
	
	public VariableResolver getVariableResolver() {
		return _vairableResolver;
	}
}
