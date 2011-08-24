


public class SS_043_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click(jq("@menu[label=\"Help\"]"));
    	waitResponse();
    	click(jq("$forum"));
    	waitResponse();
    	
    	// TODO How to verify? Can't access other window or browser tab
    }
}
