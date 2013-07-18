

public class SS_003_Test extends SSAbstractTestCase {

	/**
	 * Open file dialog
	 */
    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_openFileDialog"));
    	
    	click("$fileMenu");
    	waitResponse();
    	click("$openFile");
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_openFileDialog"));
    }
}
