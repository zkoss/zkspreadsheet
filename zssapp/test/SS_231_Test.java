import org.zkoss.ztl.JQuery;


//use mouse to auto fill
//input Jan in cell k13, Feb in cell k14
//
public class SS_231_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(10, 12, 10, 12);
		type(jq("$formulaEditor"), "Jan");
		waitResponse();
		selectCells(10, 13, 10, 13);
		type(jq("$formulaEditor"), "Feb");
		waitResponse();
		selectCells(10, 12, 10, 13);
		
		mouseOver(jq(".zsseldot"));
		waitResponse();
		mouseMove(jq(".zsseldot"));
		waitResponse();
		mouseDown(jq(".zsseldot"));
		waitResponse();
		mouseMoveAt(getSpecifiedCell(11, 17),"-2,-2");
		waitResponse();
		mouseUpAt(getSpecifiedCell(11, 17),"-2,-2");
		waitResponse();
		
		String k15value = getSpecifiedCell(10,14).text();
		String k16value = getSpecifiedCell(10,15).text();
		String k17value = getSpecifiedCell(10,16).text();
		verifyEquals(k15value,"Mar");
		verifyEquals(k16value,"Apr");
		verifyEquals(k17value,"May");
	}
}



