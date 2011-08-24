
public class SS_053_Test extends SSAbstractTestCase {

	/**
	 * Paste (formula only)
	 */
    @Override
    protected void executeTest() {
    	
    	focusOnCell(5, 14);
    	String srcCellBackgroundColor = getCellBackgroundColor(5, 14);
    	click("$copyBtn");
    	verifyTrue("Source cell shall has background for testing", srcCellBackgroundColor.length() > 0);
    	
    	//paste formula
    	focusOnCell(12, 13);
    	clickDropdownButtonMenu("$fastIconBtn $pasteDropdownBtn div.z-dpbutton-btn", "Formulas");
    	
    	focusOnCell(12, 13);
    	String formulaBarValue = jq("$formulaEditor").val();
    	String targetCellBackgroundColor = getCellBackgroundColor(12, 13);
        /**
         * Expect: 
         * paste text and formula etc.. but do not paste style
         */
    	verifyTrue("cell shall be formula", formulaBarValue.startsWith("="));
        verifyNotEquals("paste formula do not paste cell style", srcCellBackgroundColor, targetCellBackgroundColor);
    }

}
