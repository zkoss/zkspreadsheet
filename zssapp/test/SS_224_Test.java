import org.zkoss.ztl.JQuery;


//focus in a cell k13,
//press enter see if the focus goes to the cell below it, k14

public class SS_224_Test extends SSAbstractTestCase {
	@Override
	protected void executeTest() {
		selectCells(10, 12,10,12);
		
		String beforeTop = jq(".zsselect").css("top");
		String beforeLeft = jq(".zsselect").css("left");
		
		JQuery k13 = getSpecifiedCell(10, 12);
		keyDown(k13,ENTER);
		waitResponse();
		keyUp(k13,ENTER);
		waitResponse();
						
		String afterTop = jq(".zsselect").css("top");
		String afterLeft = jq(".zsselect").css("left");

		verifyEquals(beforeLeft, afterLeft);
		
		int cellHeight = getSpecifiedCell(10, 12).height();
		Integer btop = Integer.valueOf(beforeTop.substring(0, 3));
		Integer atop = Integer.valueOf(afterTop.substring(0, 3));
		
		verifyTrue(atop >= btop+cellHeight);
		
	}
}



