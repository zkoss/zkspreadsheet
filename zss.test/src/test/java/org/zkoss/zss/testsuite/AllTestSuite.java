package org.zkoss.zss.testsuite;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;

@RunWith(Suite.class)
@SuiteClasses({ModelTestSuite.class, EngineTestSuite.class, APITestSuite.class,
	FormulaTestSuite.class, FormatTestSuite.class, UITestSuite.class, ZatsTestSuite.class})
public class AllTestSuite {

}
