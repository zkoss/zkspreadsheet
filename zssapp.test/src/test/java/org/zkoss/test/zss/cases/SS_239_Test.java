/* SS_239_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Feb 1, 2012 2:13:10 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Rule;
import org.junit.Test;
import org.openqa.selenium.WebDriver;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.TestingEnvironment;
import org.zkoss.test.zss.Spreadsheet;

import com.google.guiceberry.junit4.GuiceBerryRule;
import com.google.inject.Inject;
import com.google.inject.name.Names;

/**
 * @author sam
 *
 */
public class SS_239_Test {
	
	public static class ValidationEnvironment extends TestingEnvironment{
		protected void configure() {
			//Validation URL
			bind(String.class)
			.annotatedWith(Names.named("URL"))
			.toInstance("http://localhost:8088/zssapp/test/validation.zul");
			
			bind(String.class)
			.annotatedWith(Names.named("Spreadsheet Id"))
			.toInstance("spreadsheet");
			
			super.configure();
		}
	}
	
	@Rule
	public GuiceBerryRule guiceBerry = new GuiceBerryRule(ValidationEnvironment.class);
	
	@Inject
	protected Spreadsheet spreadsheet;
	
	@Inject
	WebDriver webDriver;
	
	@Inject
	JQueryFactory jqFactory;
	
	@Inject
	protected ConditionalTimeBlocker timeBlocker;
	
	/**
	 * Refer to http://tracker.zkoss.org/browse/ZSS-84
	 * 
	 * Validation button shall show on column D, from row 1 ~ row 100
	 */
	@Test
	public void validation_buttons() {
		spreadsheet.focus(1, 3);
		
		JQuery buttons = spreadsheet.getRow(1).jq$n().children(".zsdropdown");
		Assert.assertEquals(1, buttons.length());
		Assert.assertTrue(buttons.first().isVisible());
		
		spreadsheet.focus(25, 3);
		Assert.assertEquals(25, buttons.length());
		Assert.assertTrue(buttons.first().isVisible());
	}
}
