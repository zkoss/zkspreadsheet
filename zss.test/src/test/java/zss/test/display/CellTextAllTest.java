package zss.test.display;

import java.util.Arrays;
import java.util.List;

import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;


/**
 * Test case for the function "display Excel files".
 * 
 * @author Hawk
 *
 */
@RunWith(Parameterized.class)
public class CellTextAllTest extends CellTextTest{
	
	public CellTextAllTest(String page){
		super(page);
	}

	@Parameters
	public static List<Object[]> data() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
}
