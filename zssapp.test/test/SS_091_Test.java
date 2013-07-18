import org.zkoss.ztl.JQuery;


public class SS_091_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select a cell
        JQuery cell_L_13 = getSpecifiedCell(11, 12);
        clickCell(cell_L_13);
        clickCell(cell_L_13);
        verifyEquals("L13", jq("$focusPosition .z-combobox-rounded-inp").val());
    }

}
