//close button on menubar
public class SS_045_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//verify
    	verifyFalse(jq(".toolbarMask").isVisible());
    	
    	click(jq("$closeBtn"));
    	waitResponse();
    	
    	//verify
    	verifyTrue(jq(".toolbarMask").isVisible());
    	verifyFalse(jq("$closeBtn").isVisible());
    	
    	click(jq("$fileMenu"));
    	waitResponse();
    	click(jq("$openFile"));
    	waitResponse();
    	
    	verifyTrue(isWidgetVisible("$_openFileDialog"));
   
    	//TODO: add another test to open a empty sheet verify sheet opened
    	//verifyFalse(jq(".toolbarMask").isVisible());
    	//verifyTrue(jq("$closeBtn").isVisible());
    }
}
