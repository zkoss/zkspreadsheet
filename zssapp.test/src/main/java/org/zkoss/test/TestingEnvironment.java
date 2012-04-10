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

import java.net.MalformedURLException;
import java.net.URL;

import org.openqa.selenium.WebDriver;
import org.openqa.selenium.firefox.FirefoxDriver;
import org.openqa.selenium.ie.InternetExplorerDriver;
import org.openqa.selenium.remote.DesiredCapabilities;
import org.openqa.selenium.remote.RemoteWebDriver;

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
	public WebDriver provideWebDriver(Config config) {
		Config.Browser target = config.getBrowser();
		switch (target) {
		case FIREFOX:
			return new FirefoxDriver();
		case IE:
			return new InternetExplorerDriver();
		case CHROME:
			try {
				return new RemoteWebDriver(new URL("http://localhost:9515"), DesiredCapabilities.chrome());
			} catch (MalformedURLException e) {
				throw new RuntimeException(e);
			}
		case OPERA:
			return new OperaDriver();
		}
		throw new RuntimeException("fail to create WebDriver");
	}
}
