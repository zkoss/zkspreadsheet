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

import com.google.guiceberry.GuiceBerryEnvMain;
import com.google.guiceberry.GuiceBerryModule;
import com.google.inject.AbstractModule;
import com.google.inject.Inject;
import com.google.inject.Provides;
import com.google.inject.Singleton;
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
		bind(GuiceBerryEnvMain.class).to(EnvironmentMain.class);
		
		install(ThrowingProviderBinder.forModule(this));
		install(new FactoryModuleBuilder().build(JQueryFactory.class));
	}

	static class EnvironmentMain implements GuiceBerryEnvMain  {
		@Inject @Named("URL")
		String url;

		@Inject
		WebDriver webDriver;
		
		public void run() {
			webDriver.get(url);
		}
	}
	
	//TODO: how to use CheckedProvider
	@Provides
	@Singleton
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
