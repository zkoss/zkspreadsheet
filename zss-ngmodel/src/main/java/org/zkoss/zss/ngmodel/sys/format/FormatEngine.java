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
public interface FormatEngine {

	FormatResult format(String format, Object value, FormatContext ctx);
	
}
