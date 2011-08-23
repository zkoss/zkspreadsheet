

public class SS_012_Test extends SSAbstractTestCase {

	/**
	 * Select cell area
	 */
    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	verifyFalse(isHighlighVisible());
    	
    	click("$editMenu");
    	waitResponse();
    	click("$copy");
    	waitResponse();
    	
    	/**
    	 * Expected:
    	 * selection area is visible
    	 */
    	verifyTrue(isHighlighVisible());
    }
}
