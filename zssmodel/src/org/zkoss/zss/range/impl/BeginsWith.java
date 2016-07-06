/* Between.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:13:53 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.impl.CellValue;
import org.zkoss.zss.model.impl.RuleInfo;

/**
 * @author henri
 * @since 3.9.0
 */
public class BeginsWith extends BaseMatch2 {
	public BeginsWith(String base) {
		super(base);
	}
	public BeginsWith(RuleInfo ruleInfo) {
		super(ruleInfo);
	}
	
	protected boolean matchString(String text, String b) {
		return text.startsWith(b);
	}

	protected boolean matchDouble(Double num, Double b) {
		return false;
	}
}
