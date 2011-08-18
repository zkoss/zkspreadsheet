

public class SS_008_Test extends SSAbstractTestCase {

	/**
	 * Testcase:
	 * Select cell(s)
	 * Click Edit menu
	 * 
	 * Expected:
	 * Selection remain focus
	 */
    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	
    	// TODO verify focus on "selection"
    	verifyTrue(isFocusOnCell(1, 5));
    }
}
