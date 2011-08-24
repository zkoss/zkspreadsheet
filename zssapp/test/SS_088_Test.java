

public class SS_088_Test extends SSAbstractTestCase {

	/**
	 * Sort Descending from toolbar
	 */
    @Override
    protected void executeTest() {
        selectCells(1, 18, 1, 21);
        
        // Click drop down button of Sort
        clickDropdownButtonMenu("$fastIconBtn  $sortDropdownBtn div.z-dpbutton-btn", "Descending");
        
        // Verify
        verifyTrue(getCellText(1, 18).startsWith("T"));
        verifyTrue(getCellText(1, 19).startsWith("O"));
        verifyTrue(getCellText(1, 20).startsWith("C"));
        verifyTrue(getCellText(1, 21).startsWith("A"));
    }

}
