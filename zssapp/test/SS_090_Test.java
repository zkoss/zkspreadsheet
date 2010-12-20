import org.zkoss.ztl.JQuery;


public class SS_090_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_O_6 = getSpecifiedCell(14, 5);
        clickCell(cell_O_6);
        clickCell(cell_O_6);
        System.out.println("Style before set grid line: " + cell_O_6.css("border-bottom"));
        
        click(jq("$gridlinesCheckbox"));
        waitResponse();
        cell_O_6 = getSpecifiedCell(14, 5);
        System.out.println("Style after set grid line: " + cell_O_6.css("border-bottom"));
        
        sleep(10000);
    }

}
