package org.zkoss.zss.test.selenium;

import static org.junit.Assert.fail;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;
import org.zkoss.zss.test.selenium.entity.ZSStyle;

/**
 * A skeleton of ZK client widget.
 * @author jumperchen
 *
 */
public class ZKClientTestCase extends ZKTestCase {
	protected int timeout;
	public SpreadsheetWidget getSpreadsheet() {
		return new SpreadsheetWidget(); 
	}

	public SpreadsheetWidget getSpreadsheet(String selector) {
		return new SpreadsheetWidget(jq(selector)); 
	}
	
	/**
	 * Returns the jQuery object of the selector
	 * @param selector the selector
	 */
	protected JQuery jq(String selector) {
		return JQuery.$(selector);
	}
	
	protected JQuery jq(ZSStyle style) {
		return JQuery.$(style.toString());
	}
	
	/**
	 * Waits for Ajax response. (excluding animation check)
	 * <p>By default the timeout time is specified in config.properties
	 * @see #waitResponse(int)
	 */
	protected void waitResponse() {
		waitResponse(timeout);
	}
	/**
	 * Waits for Ajax response according to the timeout attribute.
	 * @param timeout
	 * 
	 */
	protected void waitResponse(int timeout) {
		long s = System.currentTimeMillis();
		int i = 0;
		int ms = 200;
		if (isIE())
			ms /= 2;
		
		while (i < 2) { // make sure the command is triggered.
			while((Boolean) eval("return !!zAu.processing()")) {
				if (System.currentTimeMillis() - s > timeout) {
					fail("Test case timeout!");
					break;
				}
				i = 0;//reset
				sleep(ms);
			}
			i++;
		}
	}
	
	/**
	 * Returns whether is InternatExplorer Driver
	 */
	public boolean isIE() {
		return driver() instanceof InternetExplorerDriver;
	}
	
	public WebElement getActiveElement() {
		return driver().switchTo().activeElement();
	}
}
