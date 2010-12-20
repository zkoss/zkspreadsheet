import org.zkoss.ztl.JQuery;


public class SS_182_Test extends SSAbstractTestCase {

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
        
        // Right click target cell
        JQuery cell_L_13 = loadTargetCell();
        clickCell(cell_L_13);
        rightClickCell(cell_L_13);
        waitResponse();
        
        // Click Paste Special on the context menu
        click(jq("$pasteSpecial"));
        waitResponse();
        
        // Choose Formulas and number formats
        click(jq("$formulaWithNum input[id*=real]"));
        waitResponse();
        click(jq("$okBtn"));
        waitResponse();
        
        // Verify -- formula
        String formulaBarValue = jq("$formulaEditor").val();
        verifyEquals("Incorrect value: " + formulaBarValue, "=+K13", formulaBarValue);
        
        // Verify -- Number format
        String targetCellValue = cell_L_13.text();
        verifyEquals("Incorrect value: " + targetCellValue, "$0", targetCellValue);
    }
    
    private JQuery loadTargetCell() {
        return getSpecifiedCell(11, 12);
    }

}
