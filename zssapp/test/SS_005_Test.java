

public class SS_005_Test extends SSAbstractTestCase {

	/**
	 * Click File menu and select Import menuitem
	 * 
	 * Expected:
	 * Opens an Import dialog
	 */
    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_importFileDialog"));
    	
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$importFile a.z-menu-item-cnt')");
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_importFileDialog"));
    }
}
