/* GreaterThan.java

	Purpose:
		
	Description:
		
	History:
		May 18, 2016 3:13:53 PM, Created by henrichen

	Copyright (C) 2016 Potix Corporation. All Rights Reserved.
*/

package org.zkoss.zss.range.impl;

import java.io.Serializable;

import org.zkoss.zss.model.SCell;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.impl.CellValue;
import org.zkoss.zss.model.impl.RuleInfo;

/**
 * @author henri
 * @since 3.9.0
 */
public class GreaterThan2 extends BaseMatch2 {
	private static final long serialVersionUID = -2673045590406268437L;
	
	public GreaterThan2(Object b) {
		super(b);
	}
	public GreaterThan2(RuleInfo ruleInfo) {
		super(ruleInfo);
	}
	
	protected boolean matchString(String text, String b) {
		return text.compareTo(b) > 0;
	}

	protected boolean matchDouble(Double num, Double b) {
		return num.compareTo(b) > 0;
	}
}
