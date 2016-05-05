/* AbstractTableColumnAdv.java

	Purpose:
		
	Description:
		
	History:
		May 4, 2016 6:22:44 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.model.impl;

import java.io.Serializable;

import org.zkoss.zss.model.STableColumn;

/**
 * @author henri
 *
 */
public abstract class AbstractTableColumnAdv implements STableColumn,
		Serializable {
	/*package*/ abstract AbstractTableColumnAdv cloneTableColumn();
}
