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

	private final SBook book;
	private final SSheet sheet;
	private final SCell cell;
	private FunctionMapper functionMapper;
	private VariableResolver vairableResolver;

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
		this.book = book;
		this.sheet = sheet;
		this.cell = cell;
		EvaluationContributor contributor = book instanceof EvaluationContributorContainer? 
				((EvaluationContributorContainer)book).getEvaluationContributor():null;
		if(contributor!=null){
			functionMapper = contributor.getFunctionMaper(book);
			vairableResolver = contributor.getVariableResolver(book);
		}
	}

	public SBook getBook() {
		return book;
	}

	public SSheet getSheet() {
		return sheet;
	}

	public SCell getCell() {
		return cell;
	}
	
	public FunctionMapper getFunctionMapper() {
		return functionMapper;
	}
	
	public VariableResolver getVariableResolver() {
		return vairableResolver;
	}
}
