import org.zkoss.ztl.JQuery;


public class SS_094_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select a cell
        JQuery cell_L_13 = getCell_L13();
        JQuery formulaBar = jq("$formulaEditor");
        clickCell(cell_L_13);
        clickCell(cell_L_13);
        
        // Input value on formula bar
        // type(formulaBar, "123456"); // <-- It has issue.
        focus(formulaBar);
        keyPress(formulaBar, "1");
        keyPress(formulaBar, "2");
        keyPress(formulaBar, "3");
        keyPress(formulaBar, "4");
        keyPress(formulaBar, "5");
        keyPress(formulaBar, "6");
        keyPressEnter(formulaBar);
        
        // Verify
        cell_L_13 = getCell_L13();
        clickCell(cell_L_13);
        verifyEquals(formulaBar.val(), cell_L_13.text());
        sleep(12000);
    }

    private JQuery getCell_L13() {
        return getSpecifiedCell(11, 12);
    }

}
