/* SS_053_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 11, 2012 5:37:31 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import junit.framework.Assert;

import org.junit.Rule;
import org.junit.Test;
import org.openqa.selenium.WebDriver;
import org.zkoss.test.BugTestingEnvironment;
import org.zkoss.test.ConditionalTimeBlocker;
import org.zkoss.test.JQuery;
import org.zkoss.test.JQueryFactory;
import org.zkoss.test.zss.MouseDirector;
import org.zkoss.test.zss.Spreadsheet;
import org.zkoss.test.zss.ZSSTestCase;

import com.google.guiceberry.junit4.GuiceBerryRule;
import com.google.inject.Inject;
import com.google.inject.name.Names;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_053_Test {
	
	public static class B112 extends BugTestingEnvironment {
		protected void configure() {
			
			bind(String.class)
			.annotatedWith(Names.named("URL"))
			.toInstance("http://localhost:8088/zssapp/test/B112.zul");

			bind(String.class)
			.annotatedWith(Names.named("Spreadsheet Id"))
			.toInstance("spreadsheet");
			
			super.configure();
		}
	}
	
	@Rule
	public GuiceBerryRule guiceBerry = new GuiceBerryRule(B112.class);
	
	@Inject
	Spreadsheet spreadsheet;
	
	@Inject
	WebDriver webDriver;
	
	@Inject
	JQueryFactory jqFactory;
	
	@Inject
	MouseDirector mouseDirector;
	
	@Inject
	ConditionalTimeBlocker timeBlocker;
	
	JQuery jq(String selector) {
		return jqFactory.create("'" + selector + "'");
	}
	
	@Test
	public void B112() {
		
		int firstSheetWidgetSize = jq(".zswidgetpanel").children().length();
		Assert.assertEquals(1, firstSheetWidgetSize);
		
		JQuery secondSheet = jq(".zssheettab").eq(1);
		mouseDirector.click(secondSheet);
		timeBlocker.waitResponse();
		
		JQuery firstSheet = jq(".zssheettab").eq(0);
		mouseDirector.click(firstSheet);
		Assert.assertEquals(1, firstSheetWidgetSize);
		
		mouseDirector.click(jq("$setChartBook"));
		firstSheetWidgetSize = jq(".zswidgetpanel").children().length();
		Assert.assertEquals(9, firstSheetWidgetSize);
		
		secondSheet = jq(".zssheettab").eq(1);
		mouseDirector.click(secondSheet);
		Assert.assertEquals(9, firstSheetWidgetSize);
		
		firstSheet = jq(".zssheettab").eq(0);
		mouseDirector.click(firstSheet);
		Assert.assertEquals(9, firstSheetWidgetSize);
		
		mouseDirector.click(jq("$invalidateSpreadsheetBtn"));
		Assert.assertEquals(9, firstSheetWidgetSize);
	}
}
