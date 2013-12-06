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
public class FormulaParseContext extends AbstractContext {
	private final Ref dependent;

	private final NBook book;
	private final NSheet sheet;
	private final NCell cell;

	public FormulaParseContext(NCell cell,Ref dependent) {
		this(cell.getSheet().getBook(),cell.getSheet(),cell,dependent);
	}
	public FormulaParseContext(NSheet sheet,Ref dependent) {
		this(sheet.getBook(),sheet,null,dependent);
	}
	public FormulaParseContext(NBook book,Ref dependent) {
		this(book,null,null,dependent);
		
	}
	public FormulaParseContext(NBook book, NSheet sheet, NCell cell,
			Ref dependent) {
		this.book = book;
		this.sheet = sheet;
		this.cell = cell;
		this.dependent = dependent;
	}

	public Ref getDependent() {
		return dependent;
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
