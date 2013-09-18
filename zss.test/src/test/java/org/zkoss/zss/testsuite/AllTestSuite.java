package org.zkoss.zss.testsuite;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;

@RunWith(Suite.class)
@SuiteClasses({ UserModelTestSuite.class, EngineTestSuite.class,
		SysTestSuite.class, APITestSuite.class })
public class AllTestSuite {

}
