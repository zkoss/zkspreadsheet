import org.zkoss.ztl.JQuery;

//Toolbar>>Wrap text

public class SS_083_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_F_6 = loadCell();
        clickCell(cell_F_6);
        clickCell(cell_F_6);
        int originalWidth = cell_F_6.find("div").last().width();
        String beforeCss = cell_F_6.css("white-space");
        click(jq("$fastIconBtn $wrapTextBtn:visible"));
        System.out.println(">>>beforeCss:"+beforeCss);
        waitResponse();

        int currentWidth = cell_F_6.find("div").last().width();
        String afterCss = cell_F_6.css("white-space");
        System.out.println(">>>afterCss:"+afterCss);
        clickCell(cell_F_6);
        clickCell(cell_F_6);
        click(jq("$fastIconBtn $wrapTextBtn:visible"));
        waitResponse();
        
        cell_F_6 = loadCell();
        
        //TODO: in IE, width will be the same     
        verifyTrue("Unexcepted result: " + currentWidth, currentWidth >= originalWidth);
        //TODO: should verify the css? how strict?
        
    }

    private JQuery loadCell() {
        return getSpecifiedCell(5, 5);
    }

}
