import org.zkoss.ztl.JQuery;

//Edit>>Clear Content
//clear cell F21
public class SS_014_Test extends SSAbstractTestCase {

	/**
	 * Select cell area
	 * Click Edit menu
	 * Select clear content menuitem
	 * 
	 * Expected:
	 * Set selected area value to empty, keep original style 
	 */
    @Override
    protected void executeTest() {
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        String origValue = getCellText(cell_F_21);
        clickCell(cell_F_21);
        clickCell(cell_F_21);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$clearContent a.z-menu-item-cnt')");
    	waitResponse();
        String newValue = getCellText(cell_F_21);
    	
        verifyTrue("Original value=" + origValue + "" +
        		"\n, Cleared value=" + newValue, 
        		"".equals(newValue));
    }
    

}
