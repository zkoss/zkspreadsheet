package org.zkoss.zss;
import org.junit.extensions.cpsuite.ClasspathSuite;
import org.junit.extensions.cpsuite.ClasspathSuite.ClassnameFilters;
import org.junit.runner.RunWith;

@RunWith(ClasspathSuite.class)
@ClassnameFilters({
	"org.zkoss.poi.ss.usermodel.*Test",
	"org.zkoss.zss.api.impl.*Test",
	"org.zkoss.zss.engine.impl.*Test",
	"org.zkoss.zss.model.sys.impl.*Test"})
public class TestAll {

}
