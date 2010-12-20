


public class SS_044_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click("jq('@menu[label=\"Help\"] button.z-menu-btn')");
    	waitResponse();
    	click("jq('$smalltalk a.z-menu-item-cnt')");
    	waitResponse();
    	// TODO How to verify? Can't access other window or browser tab
    }
}
