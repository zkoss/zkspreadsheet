

public class SS_038_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click("jq('$insertMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$insertFormula a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if insert formula window twindow is opened
    	verifyTrue(widget(jq("@window[mode=\"highlighted\"][title=\"Insert Function\"] div.z-window-highlighted-header")).exists());
    }
}
