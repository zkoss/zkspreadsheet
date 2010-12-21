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
        
        // Click target cell - L17
        clickCell(getSpecifiedCell(11, 16));
        mouseOver(jq("$pasteDropdownBtn"));
        clickAt(jq("$pasteDropdownBtn"), "30,2");
        waitResponse();
        
        // Click Formulas
        click(jq("$pasteFormula"));
        waitResponse();
        
        // Verify
        String formulaBarValue = jq("$formulaEditor").val();
        verifyEquals("Incorrect value: " + formulaBarValue, "=K17", formulaBarValue);
    }

}
