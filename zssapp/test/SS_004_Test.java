import org.zkoss.ztl.JQuery;


public class SS_004_Test extends SSAbstractTestCase {

	
    @Override
    protected void executeTest() {
        JQuery cell = getSpecifiedCell(5, 20);
        clickCell(cell);
        clickCell(cell);

    	click("$fileMenu");
    	waitResponse();
    	click("$openFile");
    	waitResponse();
    	// TODO how to verify if open upload dialog
    	// click("jq('$uploadBtn input')");
    	
    	waitResponse();
    	// TODO verify if cell is still selected
    }
}
