import org.zkoss.ztl.JQuery;

//use edit menu toolbar 
//copy cell F21 to cell M10
public class SS_013_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        
        String sourceValue = getCellContent(cell_F_21);
        clickCell(cell_F_21);
        clickCell(cell_F_21);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$copy a.z-menu-item-cnt')");
    	waitResponse();
        JQuery cell_M_10 = getSpecifiedCell(12, 9);
        sleep(5000);
        clickCell(cell_M_10);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$paste a.z-menu-item-cnt')");
    	waitResponse();
        String targetValue = getCellContent(cell_M_10);
    	
        sleep(5000);
        
    	// TODO verify if pasted value is same as copied value
        verifyEquals("Copied value=" + sourceValue + ", Pasted value=" + targetValue, sourceValue, targetValue);
    }
}
