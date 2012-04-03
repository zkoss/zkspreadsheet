/* SS_048_Test.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Mar 21, 2012 12:45:29 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.cases;

import org.junit.Assert;
import org.junit.Test;
import org.openqa.selenium.Keys;
import org.openqa.selenium.WebElement;
import org.zkoss.test.JQuery;
import org.zkoss.test.zss.ZSSAppTest;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@ZSSTestCase
public class SS_048_Test extends ZSSAppTest {
	protected enum Chart {
		COLUMN_CHART("columnChart"),
		COLUMN_CHART_3D("columnChart3D"),
		LINE_CHART("lineChart"),
		LINE_CHART_3D("lineChart3D"),
		PIE_CHART("pieChart"),
		PIE_CHART_3D("pieChart3D"),
		BAR_CHART("barChart"),
		BAR_CHART_3D("barChart3D");
		
		private String val;
		private Chart(String val) {
			this.val = val;
		}
		
		public String toString() {
			return val;
		}
	}
	private void insertChartAndVerify(Chart type, int tRow, int lCol, int bRow, int rCol) {
		int widgetSize = jq(".zswidget").length();
		click(".zstab-insertPanel");
		
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		
		click(".zstbtn-" + type.toString().replace("3D", "") + " .zstbtn-arrow");
		click(".zsmenuitem-" + type.toString());
		
		Assert.assertEquals(widgetSize + 1, jq(".zswidget").length());
	}
	
	@Test
	public void insert_column_chart() {
		insertChartAndVerify(Chart.COLUMN_CHART, 11, 5, 13, 6);
	}
	
	@Test
	public void insert_column_chart_3D() {
		insertChartAndVerify(Chart.COLUMN_CHART_3D, 11, 5, 13, 6);
	}
	
	@Test
	public void insert_line_chart() {
		insertChartAndVerify(Chart.LINE_CHART, 11, 5, 13, 6);
	}
	
	@Test
	public void insert_line_chart_3D() {
		insertChartAndVerify(Chart.LINE_CHART_3D, 11, 5, 13, 6);
	}
	
	@Test
	public void insert_line_pie_chart() {
		insertChartAndVerify(Chart.PIE_CHART, 11, 5, 13, 6);
	}
	
	@Test
	public void insert_line_pie_chart_3D() {
		insertChartAndVerify(Chart.PIE_CHART_3D, 11, 5, 13, 6);
	}

	@Test
	public void insert_bar_chart() {
		insertChartAndVerify(Chart.BAR_CHART, 11, 5, 13, 6);
	}
	
	@Test
	public void insert_bar_chart_3D() {
		insertChartAndVerify(Chart.BAR_CHART_3D, 11, 5, 13, 6);
	}
	
	@Test
	public void insert_area_chart() {
		int widgetSize = jq(".zswidget").length();
		click(".zstab-insertPanel");
		
		int tRow = 11;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		
		click(".zstbtn-areaChart");
		
		Assert.assertEquals(widgetSize + 1, jq(".zswidget").length());
	}
	
	@Test
	public void insert_scatter_chart() {
		int widgetSize = jq(".zswidget").length();
		click(".zstab-insertPanel");
		
		int tRow = 11;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		
		click(".zstbtn-scatterChart");
		
		Assert.assertEquals(widgetSize + 1, jq(".zswidget").length());
	}
	
	@Test
	public void insert_Doughnut_chart() {
		int widgetSize = jq(".zswidget").length();
		click(".zstab-insertPanel");
		
		int tRow = 11;
		int lCol = 5;
		int bRow = 13;
		int rCol = 6;
		spreadsheet.setSelection(tRow, lCol, bRow, rCol);
		
		click(".zstbtn-otherChart .zstbtn-arrow");
		click(".zsmenuitem-doughnutChart");
		
		Assert.assertEquals(widgetSize + 1, jq(".zswidget").length());
	}
	
	@Test
	public void insert_hyperlink() {
		click(".zstab-insertPanel");
		spreadsheet.focus(11, 10);
		
		click(".zstbtn-hyperlink");
		
		Assert.assertTrue(isVisible("$_insertHyperlinkDialog"));
		
        WebElement inp = jq("$addrCombobox input.z-combobox-inp").getWebElement();
        inp.sendKeys("http://ja.wikipedia.org/wiki");
        inp.sendKeys(Keys.TAB);
        timeBlocker.waitUntil(browser.isIE6() || browser.isIE7() ? 3 : 2);
        click("$_insertHyperlinkDialog $okBtn");
        
        JQuery link = getCell(11, 10).jq$n("real").children().first();
        
        Assert.assertEquals("http://ja.wikipedia.org/wiki", link.attr("href"));
	}
}
