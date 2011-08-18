

public class SS_006_Test extends SSAbstractTestCase {

	/**
	 * Expected:
	 * 
	 * Opens Export to PDF dialog
	 */
    @Override
    protected void executeTest() {
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('@menu[label=\"Export\"] a.z-menu-cnt-img')");
    	waitResponse();
    	click("jq('$exportPdf a.z-menu-item-cnt')");
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_exportToPdfDialog"));
    }
}
