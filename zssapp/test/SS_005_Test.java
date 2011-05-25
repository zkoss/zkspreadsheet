

public class SS_005_Test extends SSAbstractTestCase {

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
