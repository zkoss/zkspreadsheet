import org.zkoss.ztl.JQuery;

public class SS_017_Test extends SSAbstractTestCase {

	/**
	 * Shift cells right
	 */
	@Override
	protected void executeTest() {
		JQuery J12 = getSpecifiedCell(9, 11);
		String origText = getCellText(J12);
		JQuery K12 = getSpecifiedCell(10, 11);
		String rightCellValue = getCellText(K12);
		clickCell(J12);
		clickCell(J12);
		click("$editMenu");
		waitResponse();
		mouseOver("$insert");		
		waitResponse();
		click("$shiftCellRight");
		waitResponse();
		String newValue = getCellText(J12);
		String newRightCellValue = getCellText(K12);
        
		/**
		 * Expected:
		 * 
		 * 1. shift cell value and cell style to right
	     * 2. current cell remain the same style but clear cell style
		 */
        verifyTrue("".equals(newValue));
        verifyEquals(newRightCellValue, origText);
        
	}
}
