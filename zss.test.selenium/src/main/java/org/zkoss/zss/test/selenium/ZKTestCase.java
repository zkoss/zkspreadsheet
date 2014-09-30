package org.zkoss.zss.test.selenium;

import org.openqa.selenium.By;
import org.openqa.selenium.TimeoutException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.openqa.selenium.support.ui.ExpectedCondition;
import org.openqa.selenium.support.ui.ExpectedConditions;
import org.openqa.selenium.support.ui.WebDriverWait;
import org.zkoss.zss.test.selenium.util.ConfigHelper;

/**
 * A skeleton of ZK Selenium test, which implements all of the methods of {@link Selenium}
 * interface.
 * 
 * @author sam
 * @author jumperchen
 * @author RaymondChao
 *
 */
public class ZKTestCase extends TestCaseBase {

	
	
	private static long ZK_TIMEOUT = 10000; // timeout for loading zk of a page
	
	/**
	 * get to a url and wait for ZK/ZSS selenium test js ready
	 * @param location
	 */
	public void getTo(String url){
		url = normalizeTestUrl(url);
		
		driver().get(url);
		
		//long t1 = System.currentTimeMillis();
		//while (true) {
		//	boolean r = (Boolean)((RemoteWebDriver)driver()).executeScript("return window.zk?true:false");
		//	if(r) break;
		//	if (zktimeout<=0 || System.currentTimeMillis() - t1 >= zktimeout){
		//		throw new RuntimeException("zk is not ready after "+zktimeout+"ms");
		//	}
		//	try {
		//		Thread.sleep(200);
		//	} catch (InterruptedException e) {}
		//}
		
		waitUntilProcessEnd(1000);
		
		((RemoteWebDriver)driver()).executeScript("zsstsDebug = "+Setup.isJsDebug()+";");
		String agentJs = AgentScripts.instance().getScript();
		((RemoteWebDriver)driver()).executeScript(agentJs);
	}
	
	public void waitUntil(long timeout, ExpectedCondition<?> condition) {
		 (new WebDriverWait(driver, timeout)).until(condition);
	}
	
	public void waitUntilProcessEnd(long millisecond) {
		try {
		// wait until ZK is start processing
		 (new WebDriverWait(driver, 1)).until(
				 ExpectedConditions.presenceOfElementLocated(By.id("zk_proc")));	 
		// wait after ZK is ready
		 (new WebDriverWait(driver, ZK_TIMEOUT, 200)).until(
				 ExpectedConditions.invisibilityOfElementLocated(By.id("zk_proc")));
		} catch (TimeoutException exception) {
			waitForTime(millisecond);
		}
	}
	
	private String normalizeTestUrl(String url) {
		url = url.toLowerCase();
		if(url.startsWith("http://") || url.startsWith("https://")){
			return url;
		}
		StringBuilder sb = new StringBuilder();
		//String base = Setup.getTestSiteUrlBase();
		String base = ConfigHelper.getInstance().getServer();
		sb.append(base);
		sb.append(ConfigHelper.getInstance().getContextPath());
		if(!base.endsWith("/")){
			sb.append("/");
		}
		if(url.startsWith("/")){
			url = url.substring(1);
		}
		sb.append(url);
		return sb.toString();
	}
	
	Object executeFns(String fn, Object... parms) {
		StringBuilder sb = new StringBuilder();
		sb.append("return zssts.").append(fn).append(".apply(zssts,arguments);");
		return ((RemoteWebDriver)driver()).executeScript(sb.toString(), parms);
	}
	
	public long getTripId(){
		return ((Number)executeFns("getTripId")).longValue();
	}
	
	public void waitForTrip(int trip, long timeout) {
		waitForTrip(getTripId(), trip, timeout);
	}

	/* package */ void waitForTrip(long lastTrip, int trip, long timeout) {
		if(trip<=0) return;
		long t1 = System.currentTimeMillis();
		while (true) {
			try {
				Thread.sleep(200);
			} catch (InterruptedException e) {
			}
			long tid2 = getTripId();
				
			if (Math.abs(tid2 - lastTrip) >= trip) {
				break;
			}
			if (timeout<=0 || System.currentTimeMillis() - t1 >= timeout)
				break;
		}
	}

	public WebElement jqSelectSingleIfAny(String selector){
		return (WebElement)executeFns("jqSelectSingle",selector);
	}
	
	public WebElement jqSelectSingle(String selector){
		WebElement elm = jqSelectSingleIfAny(selector);
		if(elm==null){
			throw new RuntimeException("can't find '"+selector+"'");
		}
		return elm;
	}
	
	/**
	 * Click on the sector represented dom element.
	 * this method also wait for a AU run-trip
	 * @param selector
	 */
	public void click(String selector) {
		WebElement elm = jqSelectSingle(selector);
		long lt = getTripId();
		elm.click();
		waitForTrip(lt, 1, Setup.getAuDelay());
	}
	
	/**
	 * send keycode to specific dom element.
	 * @param selector
	 * @param keyCode
	 */
	public void sendKeys(String selector, String keyCode) {
		WebElement elm = jqSelectSingle(selector);
		long lt = getTripId();
		elm.sendKeys(keyCode);
		waitForTrip(lt, 1, Setup.getAuDelay());
	}
	
	/**
	 * send keycode to specific dom element.
	 * @param selector
	 * @param keyCode
	 */
	public void sendKeys(WebElement element, String keyCode) {
		long lt = getTripId();
		element.sendKeys(keyCode);
		waitForTrip(lt, 1, Setup.getAuDelay());
	}
	
    /**
     * Causes the currently executing thread to sleep for the specified number
     * of milliseconds, subject to the precision and accuracy of system timers
     * and schedulers. The thread does not lose ownership of any monitors.
     * @param milliseconds the length of time to sleep in milliseconds.
     */
	protected void sleep(long milliseconds) {
		try {
			Thread.sleep(milliseconds);
		} catch (InterruptedException e) {
		}
	}
}
