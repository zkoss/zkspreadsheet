import org.zkoss.ztl.JQuery;

//use edit menu toolbar 
//copy cell F21 to cell M10
public class SS_013_Test extends SSAbstractTestCase {

	/**
	 * Paste selection
	 */
    @Override
    protected void executeTest() {
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        
        String sourceValue = getCellText(cell_F_21);
        clickCell(cell_F_21);
        clickCell(cell_F_21);
    	click("$editMenu");
    	waitResponse();
    	click("$copy");
    	waitResponse();
        JQuery cell_M_10 = getSpecifiedCell(12, 9);
        clickCell(cell_M_10);
    	click("$editMenu");
    	waitResponse();
    	click("$paste");
    	waitResponse();
        String targetValue = getCellText(cell_M_10);
    	        
        verifyEquals("Copied value=" + sourceValue + ", Pasted value=" + targetValue, sourceValue, targetValue);
    }
}
