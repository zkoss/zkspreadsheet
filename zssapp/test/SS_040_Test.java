

public class SS_040_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click("jq('$insertMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$insertImage a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if uploaded image is displayed properly
    }
}
