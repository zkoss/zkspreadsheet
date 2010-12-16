import org.zkoss.ztl.JQuery;


public class SS_083_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_10 = loadCell();
        int originalWidth = cell_B_10.width();
        clickCell(cell_B_10);
        clickCell(cell_B_10);
        click(jq("$wrapTextBtn"));
        waitResponse();
        
        cell_B_10 = loadCell();
        int currentWidth = cell_B_10.width();        
        verifyTrue("Unexcepted result: " + currentWidth, currentWidth < originalWidth);
        
        sleep(8000);
    }

    private JQuery loadCell() {
        return getSpecifiedCell(1, 9);
    }

}
