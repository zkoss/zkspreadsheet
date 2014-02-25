package org.zkoss.zss;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;
import org.zkoss.zss.model.*;
import org.zkoss.zss.ngmodel.ExporterTest;
import org.zkoss.zss.ngmodel.ExporterXlsTest;
import org.zkoss.zss.ngmodel.ImporterTest;
import org.zkoss.zss.ngmodel.ImporterXlsTest;

@RunWith(Suite.class)
@SuiteClasses({ ExporterTest.class, ImporterTest.class, ExporterXlsTest.class,
		ImporterXlsTest.class })
public class ImExporterTestSuite {

}
