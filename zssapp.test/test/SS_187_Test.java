import org.zkoss.ztl.JQuery;


public class SS_187_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_I_28 = getSpecifiedCell(8, 27);
        clickCell(cell_I_28);
        clickCell(cell_I_28);
        
        // Select multiple cells
        selectCells(8, 27, 9, 29);
        
        // Ctrl + C
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(C);
        waitResponse();
        keyUpNative(C);
        waitResponse();
        keyUpNative(CTRL);
        waitResponse();
        
        // Right click on target cell - M26
        rightClickCell(12, 25);
        
        // Click Paste Special on the context menu
        click(jq("$pasteSpecial"));
        waitResponse();
        
        // Check Skip blanks
        click(jq("$skipBlanks input[id*=real]"));
        waitResponse();
        click(jq("$okBtn"));
        
        // Verify
        verifyEquals("solid", getSpecifiedCell(12, 26).parent().css("border-right-style"));
    }

}
