

public class SS_056_Test extends SSAbstractTestCase {
	
	/**
	 * Paste Transpose
	 */
    @Override
    protected void executeTest() {
    	
    	selectCells(5, 11, 6, 11);
    	String srcText1 = getCellText(5, 11);
    	String srcText2 = getCellText(6, 11);
    	click("$copyBtn");
    	
    	/**
    	 * Transpose
    	 */
    	focusOnCell(12, 11);
    	clickDropdownButtonMenu("$fastIconBtn $pasteDropdownBtn", "Transpose");
    	
    	/**
    	 * Expect: Transpose
    	 */
    	String targetText1 = getCellText(12, 11);
    	String targetText2 = getCellText(12, 12);
    	verifyEquals(srcText1, targetText1);
    	verifyEquals(srcText2, targetText2);

    	//after transpose, the right hand side cell text shall remain empty
    	verifyTrue("".equals(getCellText(13, 11)));
    }
}