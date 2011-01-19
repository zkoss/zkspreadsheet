import org.zkoss.ztl.JQuery;


public class SS_088_Test extends SSAbstractTestCase {

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
        
        // Descending
        click(jq("$sortDescending"));
        waitResponse();
        
        // Verify
        verifyTrue(getSpecifiedCell(1, 18).text().startsWith("T"));
        verifyTrue(getSpecifiedCell(1, 19).text().startsWith("O"));
        verifyTrue(getSpecifiedCell(1, 20).text().startsWith("C"));
        verifyTrue(getSpecifiedCell(1, 21).text().startsWith("A"));
    }

}
