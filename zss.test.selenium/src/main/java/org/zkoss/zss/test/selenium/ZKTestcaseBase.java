package org.zkoss.zss.test.selenium;

import org.openqa.selenium.WebDriverException;
import org.openqa.selenium.WebElement;
import org.openqa.selenium.remote.RemoteWebDriver;

public class ZKTestcaseBase extends TestcaseBase{

	
	
	private long zktimeout = 10000; // timeout for loading zk of a page
	
	
	/**
	 * get to a url and wait for ZK/ZSS selenium test js ready
	 * @param location
	 */
	public void getTo(String url){
		url = normalizeTestUrl(url);
		
		driver.get(url);
		long t1 = System.currentTimeMillis();
		while (true) {
			boolean r = (Boolean)((RemoteWebDriver)driver).executeScript("return window.zk?true:false");
			if(r) break;
			if (zktimeout<=0 || System.currentTimeMillis() - t1 >= zktimeout){
				throw new RuntimeException("zk is not ready after "+zktimeout+"ms");
			}
			try {
				Thread.sleep(200);
			} catch (InterruptedException e) {}
		}
		((RemoteWebDriver)driver).executeScript("zsstsDebug = "+Setup.isJsDebug()+";");
		
		String agentJs = AgentScripts.instance().getScript();
		((RemoteWebDriver)driver).executeScript(agentJs);
		
		
	}
	
	private String normalizeTestUrl(String url) {
		url = url.toLowerCase();
		if(url.startsWith("http://") || url.startsWith("https://")){
			return url;
		}
		StringBuilder sb = new StringBuilder();
		String base = Setup.getTestSiteUrlBase(); 
		sb.append(base);
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
		return ((RemoteWebDriver)driver).executeScript(sb.toString(), parms);
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
}
