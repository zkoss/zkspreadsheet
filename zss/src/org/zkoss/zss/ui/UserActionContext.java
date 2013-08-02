/* UserActionContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/8/2 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui;

/**
 * 
 * @author dennis
 *
 */
public interface UserActionContext {

	public Spreadsheet getSpreadsheet();
	
	public Rect getSelection();
	
	public Object getData(String key);
	
	public String getCategory();
	
	public String getAction();
}
