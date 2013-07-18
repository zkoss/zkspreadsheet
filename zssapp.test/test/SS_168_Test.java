

public class SS_168_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	verifyFalse(isWidgetVisible("$_openFileDialog"));
    	
    	click(jq("$fileMenu"));
    	waitResponse();
    	click(jq("$openFile"));
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_openFileDialog"));
    }
}
