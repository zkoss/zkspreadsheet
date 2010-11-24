/* InsertHyperlinkButton.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 3:37:01 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul;

import org.zkoss.util.resource.Labels;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Toolbarbutton;

/**
 * @author Sam
 *
 */
public class InsertHyperlinkButton extends Toolbarbutton implements ZssappComponent {

	private Spreadsheet ss;
	
	public InsertHyperlinkButton() {
		setImage("~./zssapp/image/hyperlink.png");
		setTooltiptext(Labels.getLabel("hyperlink"));
	}
	
	public void onClick() {
		Executions.createComponents("~./zssapp/html/dialog/hyperlink/insertHyperlink.zul", null, Zssapps.newSpreadsheetArg(ss));
	}
	
	@Override
	public void bindSpreadsheet(Spreadsheet spreadsheet) {
		ss = spreadsheet;
		setWidgetListener("onClick", "this.$f('" + ss.getId() + "', true).focus(false);");
	}


	@Override
	public void unbindSpreadsheet() {
		// TODO Auto-generated method stub
		
	}

}
