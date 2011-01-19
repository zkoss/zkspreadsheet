import org.zkoss.ztl.JQuery;


public class SS_087_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
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
        
        // Ascending
        click(jq("$sortAscending"));
        waitResponse();
        
        // Verify
        verifyTrue(getSpecifiedCell(1, 18).text().startsWith("A"));
        verifyTrue(getSpecifiedCell(1, 19).text().startsWith("C"));
        verifyTrue(getSpecifiedCell(1, 20).text().startsWith("O"));
        verifyTrue(getSpecifiedCell(1, 21).text().startsWith("T"));
    }

}
