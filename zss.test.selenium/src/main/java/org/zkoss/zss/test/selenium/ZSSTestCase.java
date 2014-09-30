package org.zkoss.zss.test.selenium;

import org.openqa.selenium.WebElement;
import org.openqa.selenium.interactions.Actions;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.zkoss.zss.test.selenium.entity.ClientWidget;

public class ZSSTestCase extends ZKClientTestCase {
	
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
	
	public void closeMessageWindow() {
		
	}
}
