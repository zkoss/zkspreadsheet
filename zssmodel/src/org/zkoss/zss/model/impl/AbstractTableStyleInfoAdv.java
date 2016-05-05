/* AbstractTableStyleInfoAdv.java

	Purpose:
		
	Description:
		
	History:
		May 4, 2016 2:54:58 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SBook;
import org.zkoss.zss.model.STableStyleInfo;

/**
 * @author henri
 * @since 3.9.0
 */
public abstract class AbstractTableStyleInfoAdv implements STableStyleInfo, Serializable {
	//ZSS-1183
	/*package*/ abstract AbstractTableStyleInfoAdv cloneTableStyleInfo(SBook book);
}
