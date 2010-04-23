/* Size.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jun 5, 2008 4:15:10 PM, Created by henrichen
}}IS_NOTE

Copyright (C) 2008 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 2.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.model;

/**
 * Data carrier about width and height.
 * @author henrichen
 *
 */
public interface Size {
	/**
	 * Returns the width.
	 * @return the width.
	 */
	public int getWidth();
	
	/**
	 * Returns the height.
	 * @return the height.
	 */
	public int getHeight();
}
