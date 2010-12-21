import org.zkoss.ztl.JQuery;


public class SS_054_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select source cell
        JQuery cell_J_13 = getSpecifiedCell(9, 12);
        clickCell(cell_J_13);
        clickCell(cell_J_13);
        
        // Ctrl + C
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(C);
        waitResponse();
        keyUpNative(C);
        waitResponse();
        keyUpNative(CTRL);
        waitResponse();
        
        // Click target cell
        JQuery cell_L_13 = loadTargetCell();
        clickCell(cell_L_13);
        
        // Click Paste icon
        mouseOver(jq("$pasteDropdownBtn"));
        clickAt(jq("$pasteDropdownBtn"), "30,2");
        waitResponse();
        
        // Click Paste value
        click(jq("$pasteValue"));
        waitResponse();
        
        // Verify
        cell_L_13 = loadTargetCell();
        verifyEquals("Incorrect result! J13=80000, L13=" + cell_L_13.text(), "80000", cell_L_13.text());
    }
    
    private JQuery loadTargetCell() {
        return getSpecifiedCell(11, 12);
    }

}
