//close button on menubar
public class SS_045_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//verify
    	verifyFalse(jq(".toolbarMask").exists());
    	
    	click(jq("$closeBtn"));
    	waitResponse();
    	
    	//verify
    	verifyTrue(jq(".toolbarMask").isVisible());
    	verifyFalse(jq("$closeBtn").isVisible());
    	
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$openFile a.z-menu-item-cnt')");
    	waitResponse();
    	
    	verifyTrue(jq("$_openFileDialog div.z-window-highlighted-header").isVisible());
   
    	//TODO: add another test to open a empty sheet verify sheet opened
    	//verifyFalse(jq(".toolbarMask").isVisible());
    	//verifyTrue(jq("$closeBtn").isVisible());
    }
}
