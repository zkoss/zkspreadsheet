package org.zkoss.zss.zats;

import java.util.Arrays;
import java.util.List;

import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;


/**
 * Run test cases of the function "display Excel files" with multiple pages. 
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class CellBorderAllTest extends CellBorderTest{
	
	public CellBorderAllTest(String page){
		super(page);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
}
