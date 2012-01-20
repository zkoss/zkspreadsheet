/* ZSSAppTest.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 18, 2012 11:08:45 AM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss;

import org.junit.Before;
import org.junit.Rule;
import org.openqa.selenium.JavascriptExecutor;
import org.openqa.selenium.WebDriver;
import org.openqa.selenium.WebDriver.Navigation;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.TestingEnvironment;

import com.google.guiceberry.junit4.GuiceBerryRule;
import com.google.inject.Inject;
import com.google.inject.Injector;
import com.google.inject.name.Names;

/**
 * @author sam
 *
 */
public class ZSSAppTest {

	public static class ZSSAppEnvironment extends TestingEnvironment{
		protected void configure() {
			//ZSSApp URL
			bind(String.class)
			.annotatedWith(Names.named("URL"))
			.toInstance("http://10.1.3.115:8088/zssapp/");
			
			bind(String.class)
			.annotatedWith(Names.named("Spreadsheet Id"))
			.toInstance("spreadsheet");
			
			super.configure();
		}
	}
	
	@Rule
	public GuiceBerryRule guiceBerry = new GuiceBerryRule(ZSSAppEnvironment.class);
	
	@Inject
	Injector injector;
	
	@Inject
	WebDriver webDriver;
	
	@Inject
	JQueryFactory jqFactory;
	
	//TODO: inject-able ? use AssistedInject ?
	protected Navigation navigation;
	protected JavascriptExecutor javascriptExecutor;
	
	@Inject
	protected Spreadsheet spreadsheet;
	
	@Inject
	protected KeyboardDirector keyboardDirector;
	
	@Inject
	protected MouseDirector mouseDirector;
	
	@Inject
	protected ConditionalTimeBlocker timeBlocker;
	
	@Before
	public void setUp() {
		javascriptExecutor = (JavascriptExecutor) webDriver;
		navigation = webDriver.navigate();
		navigation.refresh();
		
		timeBlocker.waitResponse();
	}
}