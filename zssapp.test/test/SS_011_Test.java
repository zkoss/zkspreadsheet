

public class SS_011_Test extends SSAbstractTestCase {

	/**
	 * Cut selections (from menu)
	 */
    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	verifyFalse(isHighlighVisible());
    	
    	click("$editMenu");
    	waitResponse();
    	click("$cut");
    	waitResponse();
    	
    	/**
    	 * Expected:
	 	 * Selected area is highlight
    	 */
    	verifyTrue(isHighlighVisible());
    }
}
