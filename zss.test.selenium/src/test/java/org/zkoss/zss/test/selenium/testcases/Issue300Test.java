package org.zkoss.zss.test.selenium.testcases;

import org.junit.Test;
import org.zkoss.zss.test.selenium.ZSSTestcaseBase;

public class Issue300Test extends ZSSTestcaseBase {
	
	@Test
	public void testZSS369_XLSX() throws Exception{
		basename();
		
		getTo("/issue3/369-chart-xlsx.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
	}
	
	@Test
	public void testZSS369_XLS() throws Exception{
		basename();
		
		getTo("/issue3/369-chart-xls.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
	}
	
	@Test
	public void testZSS379() throws Exception{
		basename();
		
		getTo("/issue3/379-timeChinese.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
	}
	
	@Test
	public void testZSS381() throws Exception{
		basename();
		
		getTo("/issue3/381-merge-wrap.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		for(int i = 0; i < 3; i++) {
			click("@button:eq("+ i +")");
			waitForTime(500);//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS382() throws Exception{
		basename();
		
		getTo("/issue3/382-merge-hide.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		for(int i = 0; i < 2; i++) {
			click("@button:eq("+ i +")");
			waitForTime(500);//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS401() throws Exception{
		basename();
		
		getTo("/issue3/401-cutMerged.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
	}
	
	@Test
	public void testZSS428() throws Exception{
		basename();
		
		getTo("/issue3/428-chart-xlsx.zul");
		waitForTime(2000);
	}
	
	@Test
	public void testZSS401_2() throws Exception{
		basename();
		
		getTo("/issue3/401-mergeAnotherSheet.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
		
		click("@button:eq(2)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step3");
	}
	
	@Test
	public void testZSS404() throws Exception{
		basename();
		
		getTo("/issue3/404-freeze-insert.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
	}
	
	@Test
	public void testZSS436() throws Exception{
		basename();
		
		getTo("/issue3/436-wrongview.zul");
		waitForTime(2000);

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("wrongview");
	}
	
	@Test
	public void testZSS433() throws Exception{
		basename();
		
		getTo("/issue3/433-column-width.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
	}
	
	@Test
	public void testZSS434() throws Exception{
		basename();
		
		getTo("/issue3/434-merge-modify.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
	}
	
//	@Test
//	public void testZSS442() throws Exception{
//		basename();
//		
//		getTo("/issue3/442-shift-by-range.zul");
//		waitForTime(2000);
//		captureOrAssert("loadpage");
//
//		click("@button:eq(4)");
//		waitForTime(500);//for zss render in browser
//		captureOrAssert("workaround");
//	}
	
	@Test
	public void testZSS443() throws Exception{
		basename();
		
		getTo("/issue3/443-freeze-resize.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
	}
	
	@Test
	public void testZSS443_2() throws Exception{
		basename();
		
		getTo("/issue3/443-column-width.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
	}
	
	@Test
	public void testZSS450() throws Exception{
		basename();
		
		getTo("/issue3/450-delete-row.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
		
		click("@button:eq(2)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step3");
	}
	
	@Test
	public void testZSS451() throws Exception{
		basename();
		
		getTo("/issue3/451-delete-column.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		for(int i = 0; i <= 3835; i += 195) {
			setZSSScrollTop(i);
			waitForTime(500);
		}
		
		for(int i = 0; i < 1900; i += 95) {
			setZSSScrollLeft(i);
			waitForTime(500);
		}
		
		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		for(int i = 3835; i > 0; i -= 195) {
			setZSSScrollTop(i);
			waitForTime(500);
		}
		
	}
	
	@Test
	public void testZSS452() throws Exception{
		basename();
		
		getTo("/issue3/452-insertAtAnotherSheet.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
	}
	
	@Test
	public void testZSS457() throws Exception{
		basename();
		
		getTo("/issue3/457-emptyHyperlink.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
	}
	
	@Test
	public void testZSS464_XLS() throws Exception{
		basename();
		
		getTo("/issue3/464-manystyle-xls.zul");
		waitForTime(2000);
		captureOrAssert("before-apply-border");
		
		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("after-apply-border");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("check-corner");
	}
	
	@Test
	public void testZSS464_XLSX() throws Exception{
		basename();
		
		getTo("/issue3/464-manystyle-xlsx.zul");
		waitForTime(2000);
		captureOrAssert("before-apply-border");
		
		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("after-apply-border");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("check-corner");
	}
	
	@Test
	public void testZSS476() throws Exception{
		basename();
		
		getTo("/issue3/476-imageSize.zul");
		waitForTime(2000);

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("image_inserted");
	}
	
	@Test
	public void testZSS485() throws Exception{
		basename();
		
		getTo("/issue3/485-insert-row-copy.zul");
		waitForTime(2000);

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("reproduce");
	}
	
	@Test
	public void testZSS488() throws Exception{
		basename();
		
		getTo("/issue3/488-delete-first-column.zul");
		waitForTime(2000);

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("reproduce-11");
		
		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("reproduce-12");
		
		click("@button:eq(2)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("reproduce-13");
		
		click("@button:eq(3)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("reproduce-21");
		
		click("@button:eq(4)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("reproduce-22");
		
		click("@button:eq(5)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("reproduce-23");
	}
	
	@Test
	public void testZSS491() throws Exception{
		basename();
		
		getTo("/issue3/491-changeColumnChart.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
	}
	
	@Test
	public void testZSS492() throws Exception{
		basename();
		
		getTo("/issue3/492-formula-rename-sheet.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step1");
		
		click("@button:eq(2)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("step2");
	}
	
	@Test
	public void testZSS494() throws Exception{
		basename();
		
		getTo("/issue3/494-reorder-sheet-break-formula.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");

		click("@button:eq(1)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("MoveSheetXLS");
		
		click("@button:eq(2)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("MoveSheetToLastXLS");
		
		click("@button:eq(4)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("MoveSheetXLSX");
		
		click("@button:eq(5)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("MoveSheetToLastXLSX");
	}
	
	@Test
	public void testZSS499_XLS() throws Exception{
		basename();
		
		getTo("/issue3/499-moveChartData-xls.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		for(int i = 0; i < 3; i++) {
			click("@button:eq("+ i +")");
			waitForTime(500);//for zss render in browser
			captureOrAssert("step" + i);
		}
		
	}
	
	@Test
	public void testZSS499_XLSX() throws Exception{
		basename();
		
		getTo("/issue3/499-moveChartData.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		for(int i = 0; i < 3; i++) {
			click("@button:eq("+ i +")");
			waitForTime(500);//for zss render in browser
			captureOrAssert("step" + i);
		}
		
	}
	
	@Test
	public void testZSS500() throws Exception{
		basename();
		
		getTo("/issue3/500-hide-overflow.zul");
		waitForTime(2000);
		captureOrAssert("loadpage");
		
		for(int i = 1; i < 5; i++) {
			click("@button:eq("+ i +")");
			waitForTime(500);//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS505_XLSX() throws Exception{
		basename();
		
		getTo("/issue3/505-chartDisplay2007.zul");
		waitForTime(2000);
		
		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("Step1");
	}
	
	@Test
	public void testZSS505_XLS() throws Exception{
		basename();
		
		getTo("/issue3/505-chartDisplay.zul");
		waitForTime(2000);
		
		click("@button:eq(0)");
		waitForTime(500);//for zss render in browser
		captureOrAssert("Step1");
	}
	
	@Test
	public void testZSS515_XLS() throws Exception{
		basename();
		
		getTo("/issue3/515-deleteFreezeRow-xls.zul");
		waitForTime(2000);
		
		for(int i = 1; i < 10; i++) {
			click("@button:eq("+ i +")");
			waitForTime(500);//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS515_XLSX() throws Exception{
		basename();
		
		getTo("/issue3/515-deleteFreezeRow.zul");
		waitForTime(2000);
		
		for(int i = 1; i < 10; i++) {
			click("@button:eq("+ i +")");
			waitForTime(500);//for zss render in browser
			captureOrAssert("step" + i);
		}
	}
	
	@Test
	public void testZSS519() throws Exception{
		basename();
		
		getTo("/issue3/519-blankBackground2.zul");
		waitForTime(2000);
	}

}
