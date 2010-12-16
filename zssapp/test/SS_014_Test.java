import org.zkoss.ztl.JQuery;


public class SS_014_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        String origValue = getCellContent(cell_F_21);
        clickCell(cell_F_21);
        clickCell(cell_F_21);
    	click("jq('$editMenu button.z-menu-btn')");
    	waitResponse();
    	click("jq('$clearContent a.z-menu-item-cnt')");
    	waitResponse();
        String newValue = getCellContent(cell_F_21);
    	
    	// TODO verify if pasted value is same as copied value
        verifyTrue("Original value=" + origValue + ", Cleared value=" + newValue, "".equals(newValue));
    }
}
