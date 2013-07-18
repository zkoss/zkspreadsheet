import org.zkoss.ztl.JQuery;


public class SS_001_Test extends SSAbstractTestCase {

	/**
	 * Testcase:
	 * 
	 * 1. Select cell
	 * 2. Click File menu
	 * 
	 * Expected:
	 * Selection remain focus
	 */
    @Override
    protected void executeTest() {
    	verifyFalse(isFocusOnCell(5, 20));
    	JQuery cell = focusOnCell(5, 20);

    	click("$fileMenu");
    	verifyTrue(isFocusOnCell(cell));
    }
}
