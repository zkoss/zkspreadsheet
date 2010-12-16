import org.zkoss.ztl.JQuery;


public class SS_053_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_J_17 = getSpecifiedCell(9, 16);
        clickCell(cell_J_17);
        clickCell(cell_J_17);
        
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
        clickCell(getSpecifiedCell(11, 16));
        rightClickCell(getSpecifiedCell(11, 16));
        
        // Click Menuitem
        click(jq(".z-menu-item:eq(3)"));
        waitResponse();
        
        // Choose "Formulas" radio button
        click(jq("$formula input[id*=real]"));
        
        // Click "OK"
        click(jq("$okBtn"));
        
        // TODO - Verify
    }

}
