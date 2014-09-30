package org.zkoss.zss.test.selenium.testcases;

import org.junit.Ignore;
import org.junit.Test;
import org.openqa.selenium.By;
import org.openqa.selenium.WebElement;
import org.zkoss.zss.test.selenium.TestCaseBase;

@Ignore("for test zss.test.selenium self only")
public class SelfTest extends TestCaseBase {

	@Test
	public void testRegular() throws Exception{
		basename("self-test-regular");
		// And now use this to visit Google
		driver().get("http://www.google.com");
		
		captureOrAssert("loadpage");
		
		WebElement element = driver().findElement(By.name("q"));
		// Enter something to search for
		element.sendKeys("Cheese!");
		// Now submit the form. WebDriver will find the form for us from the
		// element
		element.submit();
		
		waitForTime(3000);
		
		captureOrAssert("submit");
		
	}
	
	@Test
	public void testFail() throws Exception{
		basename("self-test-fail");
		// And now use this to visit Google
		driver().get("http://www.google.com");
		
		capture("loadpage");
		
		WebElement element = driver().findElement(By.name("q"));
		// Enter something to search for
		element.sendKeys("Cheese!");
		// Now submit the form. WebDriver will find the form for us from the
		// element
		element.submit();
		
		//wait
		waitForTime(3000);
		
		capture("submit");
		
		
		//again
		driver().get("http://www.google.com");
		assertIt("loadpage");
		
		element = driver().findElement(By.name("q"));
		// Enter something to search for
		element.sendKeys("Google");
		// Now submit the form. WebDriver will find the form for us from the
		// element
		element.submit();
		
		waitForTime(3000);
		
		assertIt("submit");
	}
	
}
