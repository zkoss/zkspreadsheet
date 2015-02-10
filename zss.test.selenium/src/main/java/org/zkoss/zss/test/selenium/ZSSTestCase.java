package org.zkoss.zss.test.selenium;

import static org.junit.Assert.assertTrue;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.zkoss.zss.test.selenium.entity.ClientWidget;
import org.zkoss.zss.test.selenium.entity.JQuery;
import org.zkoss.zss.test.selenium.entity.SpreadsheetWidget;

public class ZSSTestCase extends ZKClientTestCase {
	
	public SpreadsheetWidget focusSheet() {
		SpreadsheetWidget spreadsheet = getSpreadsheet();
		spreadsheet.focus();
		return spreadsheet;
	}
	
	public void setZSSScrollTop(int val) {
		((RemoteWebDriver)driver()).executeScript("jq('@spreadsheet .zsscroll').scrollTop("+ val +");");
	}
	
	public void setZSSScrollLeft(int val) {
		((RemoteWebDriver)driver()).executeScript("jq('@spreadsheet .zsscroll').scrollLeft("+ val +");");
	}
	
	public void sendZSSKeys(String keyCode) {
		sendKeys("@spreadsheet .zsfocus", keyCode);
	}
	
	public void sendKeys(ClientWidget widget, String keyCode) {
		WebElement element = widget.toWebElement();
		long lt = getTripId();
		element.sendKeys(keyCode);
		waitForTrip(lt, 1, Setup.getAuDelay());
	}
	
	public void rightClick(WebElement element) {
		Actions action= new Actions(driver());
		action.contextClick(element).perform();
	}
	
	public void rightClick(ClientWidget widget) {
		Actions action = new Actions(driver());
		WebElement element = widget.toWebElement();
		action.contextClick(element).perform();
	}
	
	public void click(ClientWidget widget, int xOffset, int yOffset) {
		WebElement element = widget.toWebElement();
		click(element, xOffset, yOffset);
	}
	
	public void click(WebElement element, int xOffset, int yOffset) {
		Actions action= new Actions(driver());
		action.moveToElement(element, xOffset, yOffset).click().perform();
	}
	
	public void click(ClientWidget widget) {
		WebElement element = widget.toWebElement();
		click(element);
	}
	
	public void click(WebElement element) {
		Actions action= new Actions(driver());
		action.moveToElement(element).click().perform();
	}
	
	public void doubleClick(ClientWidget widget) {
		WebElement element = widget.toWebElement();
		doubleClick(element);
	}
	
	public void doubleClick(WebElement element) {
		Actions action= new Actions(driver());
		action.moveToElement(element).doubleClick().perform();
	}
	
	// coordString pattern: "1,2", means shift position to x = +1, y = +2.
	public void mouseMoveAt(ClientWidget locator, String coordString) {
		String[] coords = coordString.split(",");
		int x = Integer.parseInt(coords[0]);
		int y = Integer.parseInt(coords[1]);
		new Actions(driver()).moveToElement(locator.toWebElement(), x, y).perform();
	}
	
	public void mouseMoveAt(WebElement webElement, String coordString) {
		String[] coords = coordString.split(",");
		int x = Integer.parseInt(coords[0]);
		int y = Integer.parseInt(coords[1]);
		new Actions(driver()).moveToElement(webElement, x, y).perform();
	}
	
	public void dragAndDrop(WebElement from, WebElement to) {
		new Actions(driver()).dragAndDrop(from, to).perform();
	}
	
	public void closeMessageWindow() {
		
	}
	
	public SheetFunction sheetFunction() {
		return new SheetFunction();
	}
	
	public SheetItemUtil sheetItemUtil() {
		return new SheetItemUtil();
	}
	
	public class SheetFunction {
		/**
		 * go to tab nth
		 * @param nth start from 1, 1 means first tab
		 */
		public void gotoTab(int nth) {
			click(jq(".zssheettab:nth-child(" + nth + ")"));
		}
		
		public void copy() {
			click(jq(".zstbtn-copy.z-toolbarbutton"));
		}
		
		public void paste() {
			click(jq(".zstbtn-paste.z-toolbarbutton"));
		}
		
		public void mergeToggle() {
			click(jq(".zstbtn-mergeAndCenter.z-toolbarbutton"));
		}
	}
	
	public class SheetItemUtil {
		/**
		 * validate how close item1 and item2 are 
		 * @param item1
		 * @param item2
		 * @param tolerance
		 */
		public void assertCloseTo(JQuery item1, JQuery item2, int tolerance) {
			assertTrue(item2.exists());
			
			int top1 = item1.offsetTop();
			int left1 = item1.offsetLeft();
			int width = item1.width();
			
			int top2 = item2.offsetTop();
			int left2 = item2.offsetLeft();
			
			assertTrue(Math.abs(top1 - top2) < tolerance);
			assertTrue(Math.abs((left1 + width) - left2) < tolerance);
		}
	}
}
