package org.zkoss.zss.test.selenium.testsuites;
import org.junit.extensions.cpsuite.ClasspathSuite;
import org.junit.extensions.cpsuite.ClasspathSuite.ClassnameFilters;
import org.junit.runner.RunWith;

@RunWith(ClasspathSuite.class)
@ClassnameFilters({
	"org.zkoss.zss.test.selenium.testcases.*Test"})
public class AllTestSuite {

}
