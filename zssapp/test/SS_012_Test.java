import org.zkoss.ztl.JQuery;


public class SS_012_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
    	selectCells(1,5,5,8);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$copy a.z-menu-item-cnt')");
    	waitResponse();
    	// TODO verify if selected area is flashed with +++ border
    }
}
