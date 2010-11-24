/* CellStyleContext.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Nov 15, 2010 9:13:57 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.zul.ctrl;


/**
 * @author Ian Tsai
 *
 */
public interface CellStyleContext extends BaseContext{
	
	
	/**
	 * 
	 * @param aFontStyle
	 */
	public void doTargetChange(CellStyle aFontStyle);
	
	/**
	 * 
	 * @return
	 */
	public CellStyle getCellStyle();
	
	/**
	 * @param styleModification
	 */
	public void modifyStyle(StyleModification styleModification);
}
