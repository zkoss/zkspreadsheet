

public class SS_057_Test extends SSAbstractTestCase {

	/**
	 * Paste special
	 */
    @Override
    protected void executeTest() {
    	String srcText = getCellText(5, 11);
    	String srcTextEnd = getCellText(8, 13);
    	selectCells(5, 11, 8, 13);
    	click("$copyBtn");
    	
    	focusOnCell(12, 11);
        clickDropdownButtonMenu("$fastIconBtn $pasteDropdownBtn", "Paste Special");
        // shows dialog
        verifyTrue(isWidgetVisible("$_pasteSpecialDialog"));
        
        //paste all
        click("$_pasteSpecialDialog $okBtn");
        verifyFalse(isWidgetVisible("$_pasteSpecialDialog"));
        
        String targetText = getCellText(12, 11);
        String targetTextEnd = getCellText(15, 13);
        verifyEquals(srcText, targetText);
        verifyEquals(srcTextEnd, targetTextEnd);
    }
}
