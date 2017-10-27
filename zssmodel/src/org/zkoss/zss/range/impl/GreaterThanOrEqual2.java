/* GreaterThanOrEqual2.java

	Purpose:
		
	Description:
		
	History:
		July 18, 2017 3:13:53 PM, Created by rwenzel

	Copyright (C) 2017 Potix Corporation. All Rights Reserved.
*/
package org.zkoss.zss.range.impl;

import org.zkoss.zss.model.impl.RuleInfo;
import org.zkoss.zss.range.impl.LessThan2;

//ZSS-1336
/**
 * @author rwenzel
 * @since 3.9.2 
 */
public class GreaterThanOrEqual2 extends LessThan2 {
	private static final long serialVersionUID = 6286465320473123465L;

	public GreaterThanOrEqual2(Object base) {
		super(base);
	}

	public GreaterThanOrEqual2(RuleInfo ruleInfo) {
		super(ruleInfo);
	}

	@Override
	protected boolean matchString(String text, String b) {
		return !super.matchString(text, b);
	}

	@Override
	protected boolean matchDouble(Double num, Double b) {
		return !super.matchDouble(num, b);
	}
}
