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
package org.zkoss.zss.ngmodel.sys.formula;

import org.zkoss.zss.ngmodel.NBook;
import org.zkoss.zss.ngmodel.NCell;
import org.zkoss.zss.ngmodel.NSheet;
import org.zkoss.zss.ngmodel.sys.AbstractContext;
import org.zkoss.zss.ngmodel.sys.dependency.Ref;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public class FormulaEvaluationContext extends AbstractContext {

	private final NBook book;
	private final NSheet sheet;
	private final NCell cell;

	public FormulaEvaluationContext(NCell cell) {
		this(cell.getSheet().getBook(), cell.getSheet(), cell);
	}

	public FormulaEvaluationContext(NSheet sheet) {
		this(sheet.getBook(), sheet, null);
	}

	public FormulaEvaluationContext(NBook book) {
		this(book, null, null);

	}

	public FormulaEvaluationContext(NBook book, NSheet sheet, NCell cell) {
		this.book = book;
		this.sheet = sheet;
		this.cell = cell;
	}

	public NBook getBook() {
		return book;
	}

	public NSheet getSheet() {
		return sheet;
	}

	public NCell getCell() {
		return cell;
	}
}
