import org.zkoss.ztl.JQuery;


public class SS_089_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Custom sort - Just show dialog and check it exists or not.
        JQuery cell_B_19 = getSpecifiedCell(1, 18);
        clickCell(cell_B_19);
        clickCell(cell_B_19);
        
        // Select - B19 ~ I22
        selectCells(1, 18, 8, 21);
        
        // Click drop down button of Sort
        mouseOver(jq("@dropdownbutton:eq(2)"));
        waitResponse();
        clickAt(jq("@dropdownbutton:eq(2)"),"30,0");
        waitResponse();
        
        // Custom sort
        click(jq("$customSort"));
        waitResponse();
        
        // Verify
        verifyTrue(jq("$sortWin").isVisible());
    }

}
