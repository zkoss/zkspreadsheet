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
public class FormulaClearContext extends AbstractContext {

	private final SBook book;
	private final SSheet sheet;
	private final SCell cell;

	public FormulaClearContext(SCell cell) {
		this(cell.getSheet().getBook(), cell.getSheet(), cell);
	}

	public FormulaClearContext(SSheet sheet) {
		this(sheet.getBook(), sheet, null);
	}

	public FormulaClearContext(SBook book) {
		this(book, null, null);
	}

	private FormulaClearContext(SBook book, SSheet sheet, SCell cell) {
		this.book = book;
		this.sheet = sheet;
		this.cell = cell;
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
}
