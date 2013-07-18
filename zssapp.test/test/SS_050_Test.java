
public class SS_050_Test extends SSAbstractTestCase {

	/**
	 * Paste from "copy" using toolbar
	 */
    @Override
    protected void executeTest() {
    	String srcText = getCellText(5, 11);
    	String srcTextEnd = getCellText(8, 13);
    	selectCells(5, 11, 8, 13);
    	click("$copyBtn");
    	
    	//paste 
    	focusOnCell(12, 11);
    	click("$pasteDropdownBtn");
    	waitResponse();
    	final String targetText1 = getCellText(12, 11);
    	final String targetText1End = getCellText(15, 13);
    	final String targetStyle1 = getCellStyle(12, 11);
    	
    	//paste again
    	focusOnCell(12, 5);
    	click("$pasteDropdownBtn");
    	final String targetText2 = getCellText(12, 5);
    	final String targetText2End = getCellText(15, 7);
    	final String targetStyle2 = getCellStyle(12, 5);
    	
        
    	/**
    	 * Expected:
    	 * cell text and style shall be the same
    	 */
        verifyEquals("paste operation: src text is " + srcText + ", not equal to target: " + targetText1, 
        		srcText, targetText1);
        verifyEquals("paste operation: src text is " + srcTextEnd + ", not equal to target: " + targetText1End, srcTextEnd, targetText1End);
        verifyEquals("multi paste shall have same value", targetText1End, targetText2End);
        verifyEquals("multi paste shall have same value", targetText1, targetText2);
        verifyEquals("multi paste shall have same style", targetStyle1, targetStyle2);
    }

}
