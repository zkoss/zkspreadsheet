import org.zkoss.ztl.JQuery;


public class SS_083_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_9 = loadCell();
        int originalWidth = cell_B_9.width();
        clickCell(cell_B_9);
        clickCell(cell_B_9);
        click(jq("$wrapTextBtn"));
        waitResponse();
        
        cell_B_9 = loadCell();
        int currentWidth = cell_B_9.width();
        verifyTrue("Unexcepted result: " + currentWidth, currentWidth < originalWidth);
        
        sleep(8000);
    }

    private JQuery loadCell() {
        return getSpecifiedCell(1, 8);
    }

}
