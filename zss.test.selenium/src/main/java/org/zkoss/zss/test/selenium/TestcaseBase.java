package org.zkoss.zss.test.selenium;

import java.io.File;
import java.util.HashMap;
import java.util.LinkedHashMap;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;

import junit.framework.Assert;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.Before;
import org.junit.BeforeClass;
import org.openqa.selenium.Dimension;
import org.openqa.selenium.Point;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.remote.RemoteWebDriver;
import org.zkoss.zss.test.selenium.VisionAssert.FailAssertion;

public class TestcaseBase {

	static WebDriver driver;
	static Map<String,VisionAssert> helpers = new LinkedHashMap<String,VisionAssert>();
	
	VisionAssert visionAssert;
	
	@BeforeClass
	public static void TestcaseBase_setup(){
		driver = Setup.getDriver();
	}
	
	@AfterClass
	public static void TestcaseBase_teardown(){
		driver.quit();
	}
	
	@Before
	public void TestccaseBase_BeforeTest(){
		helpers.clear();
	}
	
	@After
	public void TestccaseBase_afterTest(){
		assertFinally();
	}
	
	public static WebDriver driver(){
		return driver;
	}
	
	/**
	 * Sets base name , it will use the caller's calss and method as the basename
	 */
	public void basename(){
		StackTraceElement[] elms = Thread.currentThread().getStackTrace();
		StackTraceElement caller = elms[2];
		String clzname = caller.getClassName();
		int i=clzname.lastIndexOf(".");
		if(i>=0){
			clzname = clzname.substring(i+1);
		}
		basename(clzname+"_"+caller.getMethodName());
	}
	public void basename(String baseName){
		visionAssert = helpers.get(baseName);
		if(visionAssert==null){
			helpers.put(baseName, visionAssert = new VisionAssert(driver, baseName));
		}
	}
	
	/**
	 * Capture a screenshot and do assertion if there is a old screenshot.
	 * If there is no old screenshot, it store it (to be a old one when next test)
	 * Note, it doesn't throw assert fail immediately for continue-assertion, you should call {@link #assertFinally()} at the end to dump all the fail assert result. 
	 * @param caseName
	 */
	public void captureOrAssert(String caseName){
		if(visionAssert==null){
			throw new IllegalStateException("set basename first");
		}
		visionAssert.captureOrAssert(caseName);
	}
	
	public void capture(String caseName){
		if(visionAssert==null){
			throw new IllegalStateException("set basename first");
		}
		visionAssert.capture(caseName);
	}
	
	public void assertIt(String caseName){
		if(visionAssert==null){
			throw new IllegalStateException("set basename first");
		}
		visionAssert.assertIt(caseName);
	}
	
	public void assertFinally(){
		for(VisionAssert va:helpers.values()){
			for(File f : va.getScreenshots()){
				System.out.println("New Screenshoot : "+f.getAbsolutePath());
			}
			va.cleanScreenshots();
		}
		StringBuilder fail = new StringBuilder();
		for(VisionAssert va:helpers.values()){
			for(FailAssertion fa : va.getFailAssertioins()){
				if(fa.getFile()!=null){
					System.out.println("New Result : "+fa.getFile().getAbsolutePath());
				}
				fail.append("\n").append(fa.getMessage());
			}
			va.cleanFailAssertions();
		}
		if(fail.length()>0){
			Assert.fail(fail.toString());
		}
	}
	
	protected void resizeWindow(int width,int height){
		driver.manage().window().setSize(new Dimension(width, height));
	}
	
	protected void relocateWindow(int x,int y){
		driver.manage().window().setPosition(new Point(x,y));
	}
	
	/**
	 * Wait for milliseconds
	 * @param ms
	 */
	public void waitForTime(long ms){
		try {
			Thread.sleep(ms);
		} catch (InterruptedException e) {}
	}

}
