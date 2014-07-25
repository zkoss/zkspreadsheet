package org.zkoss.zss.testsuite;

import org.junit.runner.RunWith;
import org.junit.runners.Suite;
import org.junit.runners.Suite.SuiteClasses;
import org.zkoss.zss.model.*;

@RunWith(Suite.class)
@SuiteClasses({ConcurrentTest.class, ExporterTest.class, ExporterXlsTest.class,
		FormatEngineTest.class, FormulaEvalTest.class, ImporterTest.class, ImporterXlsTest.class,
		InputEngineTest.class, IssueTest.class, ModelAdvancedIssueTest.class, ModelCopyTest.class,
		ModelFormulaTest.class, ModelPerformanceTest.class, ModelTest.class, PdfExporterTest.class,
		RangeTest.class, ValidationTest.class})
public class ModelTestSuite {

}
