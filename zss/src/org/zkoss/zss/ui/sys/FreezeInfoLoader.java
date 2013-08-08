/* FreezeInfoLoader.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		2013/8/8 , Created by dennis
}}IS_NOTE

Copyright (C) 2013 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.zss.ui.sys;


/**
 * @author dennis
 * @since 3.0.0
 */
public interface FreezeInfoLoader {

	/**
	 * return freeze info, zero base
	 * @param sheet
	 * @return index to freeze, -1 means no freeze
	 */
	public int getRowFreeze(Object sheet);
	
	/**
	 * return freeze info, zero base
	 * @param sheet
	 * @return index to freeze, -1 means no freeze
	 */
	public int getColumnFreeze(Object sheet);
}
