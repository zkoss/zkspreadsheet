
public class SS_051_Test extends SSAbstractTestCase {

	/**
	 * Paste (from cut), using toolbar button
	 */
    @Override
    protected void executeTest() {
    	String srcText = getCellText(5, 11);
    	String srcTextEnd = getCellText(8, 13);
    	selectCells(5, 11, 8, 13);
    	click("$cutBtn");

    	//paste
    	focusOnCell(12, 11);
    	click("$pasteDropdownBtn");
    	waitResponse();
    	
    	String srcText2 = getCellText(5, 11);
    	String srcTextEnd2 = getCellText(8, 13);
    	String targetText = getCellText(12, 11);
    	String targetTextEnd = getCellText(15, 13);
    	
    	/**
    	 * Expect
    	 * src text shall be empty
    	 */
    	verifyTrue("cut source, the cell text shall be empty", "".equals(srcText2));
    	verifyTrue("cut source, the cell text shall be empty", "".equals(srcTextEnd2));
    	/**
    	 * Expect
    	 * paste cell shall equal to source text
    	 */
    	verifyEquals("paste source, the cell text shall equal to src cell", srcText, targetText);
    	verifyEquals("paste source, the cell text shall equal to src cell", srcTextEnd, targetTextEnd);
    	
    	//paste again shall be no effect
    	focusOnCell(12, 5);
    	String text2 = getCellText(12, 5);
    	click("$pasteDropdownBtn");
    	/**
    	 * Paste from cut, it shall have no effect
    	 */
    	verifyTrue("paste from cut, the second paste shall be no effect", text2.equals(getCellText(12, 5)));
    }

}
