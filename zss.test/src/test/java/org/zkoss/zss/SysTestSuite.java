package org.zkoss.zss;
import org.junit.extensions.cpsuite.ClasspathSuite;
import org.junit.extensions.cpsuite.ClasspathSuite.ClassnameFilters;
import org.junit.runner.RunWith;

@RunWith(ClasspathSuite.class)
@ClassnameFilters({
	"org.zkoss.zss.model.sys.impl.*Test"})
public class SysTestSuite {

}
