package org.zkoss.zss;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;
import org.zkoss.zss.model.*;

@RunWith(Suite.class)
@SuiteClasses({ ExporterTest.class, ImporterTest.class, ExporterXlsTest.class,
		ImporterXlsTest.class })
public class ImExporterTestSuite {

}
