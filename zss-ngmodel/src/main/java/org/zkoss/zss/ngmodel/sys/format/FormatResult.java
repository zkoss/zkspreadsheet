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
package org.zkoss.zss.ngmodel.sys.format;
/**
 * 
 * @author dennis
 * @since 3.5.0
 */
import org.zkoss.zss.ngmodel.NColor;
import org.zkoss.zss.ngmodel.NRichText;

public interface FormatResult {

	boolean isRichText();
	
	NRichText getRichText();
	
	String getText();
	
	NColor getColor();
	
//	boolean getConditionApplied();
	
//	String getHtml();
	
}
