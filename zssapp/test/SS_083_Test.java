import org.zkoss.ztl.JQuery;


public class SS_083_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_F_6 = loadCell();
        clickCell(cell_F_6);
        clickCell(cell_F_6);
        int originalWidth = cell_F_6.find("div").last().width();
        click(jq("$fastIconBtn $wrapTextBtn:visible"));
        waitResponse();

        int currentWidth = cell_F_6.find("div").last().width();
        clickCell(cell_F_6);
        clickCell(cell_F_6);
        click(jq("$fastIconBtn $wrapTextBtn:visible"));
        waitResponse();
        
        cell_F_6 = loadCell();
        //TODO: in IE, width will be the same
        verifyTrue("Unexcepted result: " + currentWidth, currentWidth > originalWidth);
    }

    private JQuery loadCell() {
        return getSpecifiedCell(5, 5);
    }

}
