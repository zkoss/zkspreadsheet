import org.zkoss.ztl.JQuery;


public class SS_001_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell = getSpecifiedCell(5, 20);
        clickCell(cell);
        clickCell(cell);

    	click("jq('$fileMenu button.z-menu-btn')");
    	waitResponse();
    	// TODO verify if cell is still selected
    }
}
