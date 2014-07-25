package org.zkoss.zss.testsuite;

import org.junit.Ignore;
import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;
import org.zkoss.zss.zats.*;

@Ignore("need to check cases") //FIXME
@RunWith(Suite.class)
@SuiteClasses({CellBorderAllTest.class, CellDataAllTest.class, CellReferenceTest.class,
	CellTextAllTest.class, ExportTest.class, FormulaAllTest.class, FreezeHideTest.class,
	SpreadsheetAgentTest.class})
public class ZatsTestSuite {

}
