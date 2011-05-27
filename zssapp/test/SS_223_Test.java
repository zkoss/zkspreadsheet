import org.zkoss.ztl.JQuery;


//use mouse to drag edit
//move f13:g14 to k13:L14
public class SS_223_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(5, 12, 6, 13);
		
		mouseOver(jq(".zsselect"));
		waitResponse();
		mouseMove(jq(".zsselect"));
		waitResponse();
		mouseDown(jq(".zsselect"));
		waitResponse();
		mouseMoveAt(getSpecifiedCell(10,12),"2,2");
		waitResponse();
		mouseUpAt(getSpecifiedCell(10,12),"2,2");
		waitResponse();
		
		String k13value = getSpecifiedCell(10,12).text();
		String k14value = getSpecifiedCell(10,13).text();
		String l13value = getSpecifiedCell(11,12).text();
		String l14value = getSpecifiedCell(11,13).text();
		
		verifyEquals(k13value,"45,000");
		verifyEquals(k14value,"80,000");
		verifyEquals(l13value,"46,000");
		verifyEquals(l14value,"80,000");
	}
}



