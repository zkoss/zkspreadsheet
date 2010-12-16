import org.zkoss.ztl.JQuery;


public class SS_011_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$cut a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if selection area is flashed with +++ border zshighlight & zshighlight2 css class
//    	verifyTrue(widget(jq("div.z-menu-popup")).exists());

    }
}
