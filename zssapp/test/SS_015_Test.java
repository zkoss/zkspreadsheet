import org.zkoss.ztl.JQuery;


public class SS_015_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_F_4 = getSpecifiedCell(5, 4);
        clickCell(cell_F_4);
        clickCell(cell_F_4);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$clearStyle a.z-menu-item-cnt')");
    	waitResponse();
    	
    	// TODO verify if pasted cell style is cleared
    }
}
