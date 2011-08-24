

public class SS_040_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click(jq("$insertMenu"));
    	waitResponse();
    	click(jq("$insertImage"));
    	waitResponse();
    	
    	// TODO verify if uploaded image is displayed properly
    }
}
