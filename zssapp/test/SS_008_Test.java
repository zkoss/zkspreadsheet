import org.zkoss.ztl.JQuery;


public class SS_008_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	
    	// TODO verify if selection is still in focus
//    	verifyTrue(widget(jq("div.z-menu-popup")).exists());

    }
}
