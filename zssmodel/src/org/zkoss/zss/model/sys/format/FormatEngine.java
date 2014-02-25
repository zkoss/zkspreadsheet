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
package org.zkoss.zss.model.sys.format;

import org.zkoss.zss.model.SCell;

/**
 * 
 * @author dennis
 * @since 3.5.0
 */
public interface FormatEngine {

	String getEditText(SCell cell,FormatContext ctx);
	String getFormat(SCell cell, FormatContext ctx);
	FormatResult format(SCell cell, FormatContext ctx);
	FormatResult format(String format, Object value, FormatContext ctx);
	
}
