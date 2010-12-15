import org.zkoss.ztl.JQuery;


public class SS_005_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$importFile a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if open file window is opened
    	verifyTrue(widget(jq("@window[mode=\"highlighted\"][title=\"Import\"] div.z-window-highlighted-header")).exists());
    }
}
