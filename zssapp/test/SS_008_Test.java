

public class SS_008_Test extends SSAbstractTestCase {

	/**
	 * Edit menu
	 */
    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	click("$editMenu");
    	waitResponse();
    	
    	// TODO verify focus insdie "selection"
    	/**
    	 * Expected:
    	 * 
	 	 * Selection remain focus
    	 */
    	verifyTrue(isFocusOnCell(1, 5));
    }
}
