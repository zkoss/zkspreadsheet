/* StyleModification.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 16, 2010 5:57:32 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;

/**
 * @author Ian Tsai
 *
 */
public interface StyleModification {

	/**
	 * 
	 * @param style
	 * @param candidteEvt
	 */
	public void modify(CellStyle style, CellStyleContextEvent candidteEvt);
}
