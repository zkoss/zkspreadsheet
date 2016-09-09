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

import java.io.Serializable;

import org.zkoss.xel.FunctionMapper;
import org.zkoss.xel.VariableResolver;
import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SSheet;
import org.zkoss.zss.model.sys.AbstractContext;
import org.zkoss.zss.model.sys.dependency.Ref;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class FormulaEvaluationContext extends AbstractContext implements Serializable {
	private static final long serialVersionUID = 2411072362379525686L;
	
	private final SBook _book;
	private final SSheet _sheet;
	private final SCell _cell;
	private FunctionMapper _functionMapper;
	private VariableResolver _vairableResolver;
	private final Ref _dependent;
	private final int[] _offset; //ZSS-1142: [rowOffset, colOffset] for ConditionalFormatting relative formula evaluation
	private final boolean _externFormula; //ZSS-1271: indicate a ConditionalFormatting formula evaluation(formula not defined inside the cell) 

	public FormulaEvaluationContext(SCell cell,Ref dependent) {
		this(cell.getSheet().getBook(), cell.getSheet(), cell,dependent, null, false); //ZSS-1271
	}

	public FormulaEvaluationContext(SSheet sheet,Ref dependent) {
		this(sheet.getBook(), sheet, null,dependent, null, false);//ZSS-1271
	}

	public FormulaEvaluationContext(SBook book,Ref dependent) {
		this(book, null, null,dependent, null, false);//ZSS-1271
	}

	//ZSS-1142
	public FormulaEvaluationContext(SSheet sheet,Ref dependent, int[] offset) {
		this(sheet.getBook(), sheet, null,dependent, offset, true);//ZSS-1271
	}
	//ZSS-1257
	public FormulaEvaluationContext(SSheet sheet,SCell cell, Ref dependent, int[] offset, boolean externFormula) {
		this(sheet.getBook(), sheet, cell,dependent, offset, externFormula);//ZSS-1271
	}

	private FormulaEvaluationContext(SBook book, SSheet sheet, SCell cell,Ref dependent, int[] offset, boolean externformula) { //ZSS-1142, ZSS-1271
		this._book = book;
		this._sheet = sheet;
		this._cell = cell;
		this._dependent = dependent;
		EvaluationContributor contributor = book instanceof EvaluationContributorContainer? 
				((EvaluationContributorContainer)book).getEvaluationContributor():null;
		if(contributor!=null){
			_functionMapper = contributor.getFunctionMaper(book);
			_vairableResolver = contributor.getVariableResolver(book);
		}
		this._offset = offset; //ZSS-1142
		this._externFormula = externformula; //ZSS-1271
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
	public Ref getDependent() {
		return _dependent;
	}	
	
	public FunctionMapper getFunctionMapper() {
		return _functionMapper;
	}
	
	public VariableResolver getVariableResolver() {
		return _vairableResolver;
	}
	
	//ZSS-1142
	public int[] getOffset() {
		return _offset;
	}
	
	//ZSS-1271
	public boolean isExternalFormula() {
		return _externFormula;
	}
}
