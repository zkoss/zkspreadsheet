/* LessThanOrEqual2.java

	Purpose:
		
	Description:
		
	History:
		July 18, 2017 3:13:53 PM, Created by rwenzel

	Copyright (C) 2017 Potix Corporation. All Rights Reserved.
*/package org.zkoss.zss.range.impl;

import org.zkoss.zss.model.impl.RuleInfo;
import org.zkoss.zss.range.impl.GreaterThan2;

//ZSS-1336
/**
 * @author rwenzel
 * @since 3.9.2
 */
public class LessThanOrEqual2 extends GreaterThan2 {
	private static final long serialVersionUID = -6314164696828988128L;

	public LessThanOrEqual2(Object base) {
		super(base);
	}

	public LessThanOrEqual2(RuleInfo ruleInfo) {
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
