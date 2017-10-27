/* NotContainsText.java

	Purpose:
		
	Description:
		
	History:
		July 18, 2017 3:13:53 PM, Created by rwenzel

	Copyright (C) 2017 Potix Corporation. All Rights Reserved.
*/
package org.zkoss.zss.range.impl;

import org.zkoss.zss.model.impl.RuleInfo;
import org.zkoss.zss.range.impl.ContainsText;

//ZSS-1336
/**
 * @author rwenzel
 * @since 3.9.2
 */
public class NotContainsText extends ContainsText {
	private static final long serialVersionUID = -5872303750074732584L;

	public NotContainsText(String base) {
		super(base);
	}
	
	public NotContainsText(RuleInfo ruleInfo) {
		super(ruleInfo);
	}
	
	protected boolean matchString(String text, String b) {
		return !super.matchString(text, b);
	}

	protected boolean matchDouble(Double num, Double b) {
		return false;
	}
	
}
