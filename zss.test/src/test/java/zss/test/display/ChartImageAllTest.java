package zss.test.display;

import java.util.Arrays;
import java.util.List;

import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

@RunWith(Parameterized.class)
public class ChartImageAllTest extends ChartImageTest {

	public ChartImageAllTest(String page) {
		super(page);
	}
	
	@Parameters
	public static List<Object[]> getTestPage() {
		Object[][] data = new Object[][] { { "/display.zul" }, { "/display2003.zul"}};
		return Arrays.asList(data);
	}
}
