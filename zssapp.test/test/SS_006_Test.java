

public class SS_006_Test extends SSAbstractTestCase {


	/**
	 * Open export PDF dialog
	 */
    @Override
    protected void executeTest() {
    	click("$fileMenu");
    	mouseOver(jq("@menu[label=\"Export\"]"));
    	click("$exportPdf");
    	waitResponse();
    	
    	/**
    	 * Expected:
    	 * Opens Export to PDF dialog
    	 */
    	verifyTrue("Export PDF dialog shall be visible", isWidgetVisible("$_exportToPdfDialog"));
    }
}
