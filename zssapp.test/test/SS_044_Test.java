


public class SS_044_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	click(jq("@menu[label=\"Help\"]"));
    	waitResponse();
    	click(jq("$smalltalk"));
    	waitResponse();
    	// TODO How to verify? Can't access other window or browser tab
    }
}
