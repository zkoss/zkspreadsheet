package zss.test;

import java.util.Arrays;
import java.util.List;

import org.junit.runner.RunWith;
import org.junit.runners.Parameterized;
import org.junit.runners.Parameterized.Parameters;

@RunWith(Parameterized.class)
public class FormulaAllTest extends FormulaTest {

	public FormulaAllTest(String testPage){
		super(testPage);
	}
	
	@Parameters
	public static List<Object[]> getTestPages() {
		Object[][] data = new Object[][] { { "/formula.zul" }, { "/formula2003.zul"}};
		return Arrays.asList(data);
	}
}
