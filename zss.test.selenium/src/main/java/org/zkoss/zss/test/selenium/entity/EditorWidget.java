/* Editor.java

	Purpose:
		
	Description:
		
	History:
		Mon, Jul 21, 2014  2:16:39 PM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.TestCaseBase;

/**
 * 
 * @author RaymondChao
 */
public class EditorWidget extends Widget {
	
	public enum EditorType {
	    INLINE,
	    FORMULABAR
	}
	
	private static final String EDITOR = ".dp.getEditor(%1$s)";
	
	public EditorWidget(SheetCtrlWidget spreadsheet) {
		setSelector(spreadsheet.getSelector().toString() + String.format(EDITOR, ""));
	}
	
	public EditorWidget(SheetCtrlWidget spreadsheet, EditorType type) {
		setSelector(spreadsheet.getSelector().toString() +
				String.format(EDITOR, type == EditorType.FORMULABAR ? "'formulabarEditing'" : ""));
	}
	
	public WebElement toWebElement() {
		return (WebElement) TestCaseBase.eval(this.getResult("getTextNode"));
	}
	
	public String getValue() {
		return (String) get("value");
	}
}
