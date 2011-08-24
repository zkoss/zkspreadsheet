import org.zkoss.ztl.JQuery;


public class SS_025_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell = getSpecifiedCell(5, 20);
        clickCell(cell);
        clickCell(cell);

    	click("$viewMenu");
    	waitResponse();
    	// TODO verify if cell is still selected. Need to check x,y coordinates instead of display
    	String display = jq("div.zsselect").css("display");
    	verifyTrue("css display=" + display, !"none".equals(display));
    }
}
