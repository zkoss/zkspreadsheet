/* CellVisitor.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Sep 17, 2010 2:41:18 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
	This program is distributed under GPL Version 3.0 in the hope that
	it will be useful, but WITHOUT ANY WARRANTY.
}}IS_RIGHT
*/
package org.zkoss.zss.ui.impl;

/**
 * @author Sam
 *
 */
public interface CellVisitor {

	/**
	 * @param context
	 */
	void handle(CellVisitorContext context);

}
