import org.zkoss.ztl.JQuery;


public class SS_056_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select source cell
    	selectCells(7, 12, 8, 12);
        
        // Ctrl + C
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(C);
        waitResponse();
        keyUpNative(C);
        waitResponse();
        keyUpNative(CTRL);
        waitResponse();
        
        // Right click target cell
        JQuery cell_L_13 = getSpecifiedCell(11, 12);
        clickCell(cell_L_13);

        // Click Paste icon
        mouseOver(jq("$pasteDropdownBtn"));
        clickAt(jq("$pasteDropdownBtn"), "30,2");
        waitResponse();
        
        // Click Transpose
        click(jq("$pasteTranspose"));
        waitResponse();
        
        // Verify
        String val1 = getSpecifiedCell(11, 12).text();
        String val2 = getSpecifiedCell(11, 13).text();
        verifyEquals("46,500", val1);
        verifyEquals("56,000", val2);
    }
}