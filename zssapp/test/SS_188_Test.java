import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_188_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_I_28 = getSpecifiedCell(8, 27);
        clickCell(cell_I_28);
        clickCell(cell_I_28);
        
        // Select multiple cells
        selectCells(8, 11, 9, 12);
        
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
        click(jq("$transpose input[id*=real]"));
        waitResponse();
        click(jq("$okBtn"));
        
        // Verify
        verifyTrue(ColorVerifingHelper.isEqualColor(
                getSpecifiedCell(12, 25).parent().css("background-color"), 
                getSpecifiedCell(12, 26).parent().css("background-color")));
    }

}
