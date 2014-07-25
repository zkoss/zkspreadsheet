package org.zkoss.zss.test.selenium;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.zkoss.zss.test.entity.ClientWidget;

public class ZSSTestCaseBase extends ZKTestCaseBase {
	
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
		WebElement element = getWebElement(widget);
		long lt = getTripId();
		element.sendKeys(keyCode);
		waitForTrip(lt, 1, Setup.getAuDelay());
	}

	public void click(int row, int col) {
		Actions action= new Actions(driver());
		WebElement element = (WebElement) executeFns("getCell", row, col);
		if (element == null) {
			throw new RuntimeException("can't find cell(" + row + ", " + col + ")");
		}
		action.contextClick(element).build().perform();
	}
	
	public void rightClick(WebElement element) {
		Actions action= new Actions(driver());
		action.contextClick(element).perform();
	}
	
	public void rightClick(ClientWidget widget) {
		Actions action = new Actions(driver());
		WebElement element = getWebElement(widget);
		action.contextClick(element).perform();
	}
	
	public void click(WebElement element, int xOffset, int yOffset) {
		Actions action= new Actions(driver());
		action.moveToElement(element, xOffset, yOffset).click().perform();
	}
	
	public void click(String selector, int xOffset, int yOffset) {
		WebElement element = jqSelectSingle(selector);
		Actions action= new Actions(driver());
		action.moveToElement(element, xOffset, yOffset).click().perform();
	}
	
	public void click(ClientWidget widget) {
		Actions action = new Actions(driver());
		WebElement element = getWebElement(widget);
		if (element == null) {
			throw new RuntimeException("can't find element");
		}
		action.moveToElement(element).click().build().perform();
	}
}
