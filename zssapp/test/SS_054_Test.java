import org.zkoss.ztl.JQuery;


public class SS_054_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select source cell
        JQuery cell_I_13 = getSpecifiedCell(8, 12);
        String origValue = numberOnly(getCellText(cell_I_13));
        clickCell(cell_I_13);
        clickCell(cell_I_13);
        
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
        
        //TODO: strange, in IE8, this line works like click "paste", not "paste value"?
        // Click Paste value
        click(jq("$pasteValue"));
        waitResponse();
        
        // Verify
        cell_L_13 = getSpecifiedCell(11, 12);
        String result = cell_L_13.text();
        System.out.println(">>>result "+result);
        verifyEquals("Incorrect result! L13=" + result, "56000", result);
        sleep(5000);
    }
    
    private JQuery loadTargetCell() {
        return getSpecifiedCell(11, 12);
    }
}
