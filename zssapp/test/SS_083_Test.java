import org.zkoss.ztl.JQuery;


public class SS_083_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_F_6 = loadCell();
        clickCell(cell_F_6);
        clickCell(cell_F_6);
        click(jq("$wrapTextBtn"));
        waitResponse();
        
        int originalWidth = cell_F_6.width();
        clickCell(cell_F_6);
        clickCell(cell_F_6);
        click(jq("$wrapTextBtn"));
        waitResponse();
        
        cell_F_6 = loadCell();
        int currentWidth = cell_F_6.width();
        verifyTrue("Unexcepted result: " + currentWidth, currentWidth < originalWidth);
        
        sleep(8000);
    }

    private JQuery loadCell() {
        return getSpecifiedCell(5, 5);
    }

}
