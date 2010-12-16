import org.zkoss.ztl.JQuery;


public class SS_064_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        click(jq("$underlineBtn"));
        
        cell_B_8 = getSpecifiedCell(1, 7);
        String style = cell_B_8.css("text-decoration");
        
        if (style != null) {
            verifyTrue("Unexcepted result: " + style, "underline".equalsIgnoreCase(style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
        
        sleep(4000);
    }

}
