package zss.test;

import org.junit.After;
import org.junit.AfterClass;
import org.junit.BeforeClass;
import org.zkoss.zats.mimic.Zats;

public class SpreadsheetTestCaseBase {
	@BeforeClass
	public static void init() {
		Zats.init("./src/main/webapp"); 
	}

	@AfterClass
	public static void end() {
		Zats.end();
	}

	@After
	public void after() {
		Zats.cleanup();
	}

}
