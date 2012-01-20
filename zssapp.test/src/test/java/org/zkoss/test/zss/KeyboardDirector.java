/* KeyboardDirector.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 4:06:44 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.openqa.selenium.Keys;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebElement;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JavascriptActions;

import com.google.inject.Inject;

/**
 * @author sam
 *
 */
public class KeyboardDirector {
	
	final WebDriver webDriver;
	
	final Spreadsheet spreadsheet;
	
	final ConditionalTimeBlocker timeBlocker;
	
	@Inject
	/*package*/ KeyboardDirector (Spreadsheet spreadsheet, ConditionalTimeBlocker timeBlocker, WebDriver webDriver) {
		this.spreadsheet = spreadsheet;
		this.timeBlocker = timeBlocker;
		this.webDriver = webDriver;
	}
	
	public void sendKeys(int row, int col, CharSequence keys) {
		spreadsheet.focus(row, col);
		
		//TODO: this assume spreadsheet on focus mode, need to deal with editing mode ?
		final CharSequence first = keys.subSequence(0, 1);
		WebElement webElement = spreadsheet.jq$focus().getWebElement();
		webElement.sendKeys(first);
		timeBlocker.waitResponse();

		if (keys.length() > 1) {
			final CharSequence rest = keys.subSequence(1, keys.length());

			webElement = spreadsheet.getInlineEditor().getWebElement();
			webElement.sendKeys(rest);
			timeBlocker.waitResponse();
		}
	}
	
	public void setEditText(int row, int col, CharSequence keys) {

		sendKeys(row, col, keys);
		spreadsheet.getInlineEditor().jq$n().getWebElement().sendKeys(Keys.ENTER);
//		new JavascriptActions(webDriver).enter(spreadsheet.getInlineEditor().jq$n()).perform();
		
		timeBlocker.waitResponse();
	}
	
	public void enter(JQuery target) {
		if (target == null) {
			target = spreadsheet.jq$focus();
		}
		new JavascriptActions(webDriver).enter(target).perform();
		timeBlocker.waitResponse();
	}
	
	public void esc(JQuery target) {
		if (target == null) {
			target = spreadsheet.jq$focus();
		}
		new JavascriptActions(webDriver).esc(target).perform();
		timeBlocker.waitResponse();
	}
	
	public void delete(int row, int col) {
		spreadsheet.focus(row, col);
		JQuery target = spreadsheet.jq$n();
		new JavascriptActions(webDriver).delete(target).perform();
		timeBlocker.waitResponse();
	}
}
