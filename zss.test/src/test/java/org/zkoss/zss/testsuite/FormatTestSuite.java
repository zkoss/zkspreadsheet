package org.zkoss.zss.testsuite;
import org.junit.extensions.cpsuite.ClasspathSuite;
import org.junit.extensions.cpsuite.ClasspathSuite.ClassnameFilters;
import org.junit.runner.RunWith;

@RunWith(ClasspathSuite.class)
@ClassnameFilters({
	"org.zkoss.zss.api.impl.format.*Test"})
public class FormatTestSuite {

}
