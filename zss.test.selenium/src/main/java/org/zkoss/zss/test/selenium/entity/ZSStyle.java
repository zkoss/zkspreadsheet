/* ZSStyle.java

	Purpose:
		
	Description:
		
	History:
		Wed, Aug 27, 2014  3:13:56 PM, Created by RaymondChao

Copyright (C) 2014 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.test.selenium.entity;

/**
 * @author RaymondChao
 */
public enum ZSStyle {
	WIDGET("zswidget"),
	SAVE_BOOK_BTN("zstbtn-saveBook"),
	FONT_BOLD_BTN("zstbtn-fontBold"),
	HYPERLINK_BTN("zstbtn-hyperlink"),
	ACTIVE_BTN("zstbtn-seld"),
	HORIZONTAL_ALIGN("zstbtn-horizontalAlign"),
	ALIGN_LEFT("zsmenuitem-horizontalAlignLeft"),
	ALIGN_RIGHT("zsmenuitem-horizontalAlignRight"),
	FORMULABAR_OK_BTN("zsformulabar-okbtn"),
	FONT_SIZE_BOX("zsfontsize"),
	DROPDOWN_BTN("zsdropdown"),
	STYLEPANEL_MENU("zsstylepanel-menu"),
	SHEET_TAB("zssheettab"),
	DATA("zsdata"),
	CELL("zscell"),
	MASK("zssmask"),
	MENUPOPUP_OPEN("z-menupopup-open"); //some menupops have 2 duplicate elements, select the open one
	 
	private final String name;
	 
	private ZSStyle() {
		this("");
	}

	private ZSStyle(String name) {
		this.name = name;
	}
	
	public String getName(){
		return name;
	}

	@Override
	public String toString() {
		return "." + name;
	}
}
