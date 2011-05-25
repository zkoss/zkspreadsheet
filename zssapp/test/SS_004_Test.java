import org.zkoss.ztl.JQuery;


public class SS_004_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell = getSpecifiedCell(5, 20);
        clickCell(cell);
        clickCell(cell);

    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$openFile a.z-menu-item-cnt')");
    	waitResponse();
    	// TODO how to verify if open upload dialog
    	// click("jq('$uploadBtn input')");
    	
    	waitResponse();
    	// TODO verify if cell is still selected
    }
}
