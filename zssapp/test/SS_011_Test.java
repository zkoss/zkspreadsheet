

public class SS_011_Test extends SSAbstractTestCase {

	/**
	 * Testcase:
	 * Select cell area
	 * Click Edit menu
	 * Select cut menuitem
	 * 
	 * Expected:
	 * Selected area is highlight
	 */
    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	verifyFalse(isHighlighVisible());
    	
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$cut a.z-menu-item-cnt')");
    	waitResponse();
    	
    	verifyTrue(isHighlighVisible());
    }
}
