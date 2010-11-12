/* SpreadsheetCompositeComponent.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 11, 2010 3:38:40 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.app.zul;

import org.zkoss.zss.ui.Spreadsheet;

/**
 * @author Sam
 *
 */
public interface ZssappComponent {
	
	public void setSpreadsheet(Spreadsheet spreadsheet);
	
	public Spreadsheet getSpreadsheet();
}
