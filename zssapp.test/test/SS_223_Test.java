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
		
		String k13value = getCellText(10,12);
		String k14value = getCellText(10,13);
		String l13value = getCellText(11,12);
		String l14value = getCellText(11,13);
		
		verifyEquals(k13value,"45,000");
		verifyEquals(k14value,"80,000");
		verifyEquals(l13value,"46,000");
		verifyEquals(l14value,"80,000");
	}
}



