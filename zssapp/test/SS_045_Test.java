//close button on menubar
public class SS_045_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	//verify
    	verifyFalse(jq(".toolbarMask").exists());
    	
    	click(jq(".z-toolbarbutton img:eq(1)"));
    	waitResponse();
    	
    	//verify
    	verifyTrue(jq(".toolbarMask").exists());
    	
    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$openFile a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if open file window is opened
    	verifyTrue(widget(jq("$fileOpenWin div.z-window-highlighted-header")).exists());
    }
}
