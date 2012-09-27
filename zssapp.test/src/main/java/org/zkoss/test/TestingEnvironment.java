/* SeleniumEnv.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 11:10:15 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import java.io.File;
import java.net.URISyntaxException;
import java.net.URL;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.chrome.ChromeDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;

import com.google.common.testing.TearDown;
import com.google.common.testing.TearDownAccepter;
import com.google.guiceberry.GuiceBerryModule;
import com.google.guiceberry.TestId;
import com.google.guiceberry.TestScoped;
import com.google.guiceberry.TestWrapper;
import com.google.inject.AbstractModule;
import com.google.inject.Provides;
import com.google.inject.assistedinject.FactoryModuleBuilder;
import com.google.inject.name.Named;
import com.google.inject.throwingproviders.ThrowingProviderBinder;
import com.opera.core.systems.OperaDriver;

/**
 * @author sam
 *
 */
public class TestingEnvironment extends AbstractModule {
	
	@Override
	protected void configure() {
		install(new GuiceBerryModule());
		
		install(ThrowingProviderBinder.forModule(this));
		install(new FactoryModuleBuilder().build(JQueryFactory.class));
	}
	
	@Provides
	TestWrapper getTestWrapper(final TestId testId, final TearDownAccepter tearDownAccepter, final WebDriver webDriver, final @Named("URL") String url) {
		return new TestWrapper() {

			@Override
			public void toRunBeforeTest() {
				tearDownAccepter.addTearDown(new TearDown() {
					
					@Override
					public void tearDown() throws Exception {
						//used it when @TestScoped
						webDriver.quit();
						
						//TODO: how to quit webDriver after run last unit test
						//when use @Singleton
//						webDriver.navigate().refresh();
					}
				});
				webDriver.get(url);
			}
		};
	}
	
	//TODO: how to use CheckedProvider
	@Provides
//	@Singleton //issue: how to quit webDriver
	@TestScoped //issue: slow, need to open browser each unit test
	public WebDriver provideWebDriver(Config config) throws URISyntaxException {
		Config.Browser target = config.getBrowser();
		ClassLoader contextClassLoader = Thread.currentThread().getContextClassLoader();
		URL driverURL = null;
		switch (target) {
		case FIREFOX:
			
			return new FirefoxDriver();
		case IE:
			driverURL = contextClassLoader.getResource("archive/IEDriverServer.exe");
			if (driverURL == null) {
				throw new NullPointerException("Cannot find IEDriverServer instance, copy IEDriverServer.exe from driver to /src/archive/");
			}
			System.setProperty("webdriver.ie.driver", new File(driverURL.toURI()).getAbsolutePath());
			return new InternetExplorerDriver();
		case CHROME:
			driverURL = contextClassLoader.getResource("archive/chromedriver.exe");
			if (driverURL == null) {
				throw new NullPointerException("Cannot find chromedriver instance, copy chromedriver.exe from driver to /src/archive/");
			}
			System.setProperty("webdriver.chrome.driver", new File(driverURL.toURI()).getAbsolutePath());
			return new ChromeDriver();
			/*
			try {
				return new RemoteWebDriver(new URL("http://localhost:9515"), DesiredCapabilities.chrome());
			} catch (MalformedURLException e) {
				throw new RuntimeException(e);
			}
			*/
		case OPERA:
			return new OperaDriver();
		}
		throw new RuntimeException("fail to create WebDriver");
	}
}
