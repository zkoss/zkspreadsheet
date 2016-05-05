/* AbstractTableStyleAdv.java

	Purpose:
		
	Description:
		
	History:
		May 4, 2016 2:58:49 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.STableStyle;

/**
 * @author henri
 * @since 3.9.0
 */
public abstract class AbstractTableStyleAdv implements STableStyle,
		Serializable {
	/*package*/ abstract AbstractTableStyleAdv cloneTableStyle(SBook book);
}
