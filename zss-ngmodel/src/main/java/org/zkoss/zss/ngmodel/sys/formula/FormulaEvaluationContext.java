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

import java.util.Locale;

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
public class FormulaEvaluationContext extends AbstractContext{
	
	private NBook book;
	private NSheet sheet;
	private NCell cell;
	private Ref dependent;
	
	public FormulaEvaluationContext(NBook book) {
	}
	
	public NBook getBook(){
		return book;
	}

	public NSheet getSheet() {
		return sheet;
	}

	public NCell getCell() {
		return cell;
	}

	public Ref getDependent() {
		return dependent;
	}
}
