import org.zkoss.ztl.JQuery;


//focus in a cell k15,
//input "=" in formula, and use mouse to choose cell F15
//then press enter

public class SS_225_Test extends SSAbstractTestCase {
	
	
	@Override
	protected void executeTest() {
		selectCells(10, 14, 10, 14);
		type(jq("$formulaEditor"), "=");
		waitResponse();

		selectCells(5, 14, 5, 14);
		JQuery f15 = getSpecifiedCell(5, 14);
		keyDown(jq("$formulaEditor"),ENTER);
		waitResponse();
		keyUp(jq("$formulaEditor"),ENTER);
		waitResponse();

		JQuery k15 = getSpecifiedCell(10, 14);
		String result = k15.text();
		verifyEquals(result, "125000");
	}
}



