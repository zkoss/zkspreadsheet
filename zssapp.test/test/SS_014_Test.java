import org.zkoss.ztl.JQuery;

//Edit>>Clear Content
//clear cell F21
public class SS_014_Test extends SSAbstractTestCase {

	/**
	 * Clear Content
	 */
    @Override
    protected void executeTest() {
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        String origValue = getCellText(cell_F_21);
        clickCell(cell_F_21);
        clickCell(cell_F_21);
    	click("$editMenu");
    	waitResponse();
    	click("$clearContent");
    	waitResponse();
        String newValue = getCellText(cell_F_21);
    	
        /**
         * Expected:
	     * clear selection text and style 
         */
        verifyTrue("Original value=" + origValue + "" +
        		"\n, Cleared value=" + newValue, 
        		"".equals(newValue));
    }
    

}
