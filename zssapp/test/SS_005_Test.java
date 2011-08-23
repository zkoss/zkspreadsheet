

public class SS_005_Test extends SSAbstractTestCase {

	/**
	 * Import file dialog
	 */
    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_importFileDialog"));
    	
    	click("$fileMenu");
    	waitResponse();
    	click("$importFile");
    	waitResponse();
    	
    	verifyTrue("Improt file dialog shall be visible", isWidgetVisible("$_importFileDialog"));
    }
}
