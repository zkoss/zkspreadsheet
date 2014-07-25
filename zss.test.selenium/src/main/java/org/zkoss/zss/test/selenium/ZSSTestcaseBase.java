package org.zkoss.zss.test.selenium;

import org.openqa.selenium.remote.RemoteWebDriver;

public class ZSSTestcaseBase extends ZKTestcaseBase {
	
	public void setZSSScrollTop(int val) {
		((RemoteWebDriver)driver).executeScript("jq('@spreadsheet .zsscroll').scrollTop("+ val +");");
	}
	
	public void setZSSScrollLeft(int val) {
		((RemoteWebDriver)driver).executeScript("jq('@spreadsheet .zsscroll').scrollLeft("+ val +");");
	}
	
	public void sendZSSKeys(String keyCode) {
		sendKeys("@spreadsheet .zsfocus", keyCode);
	}

}
