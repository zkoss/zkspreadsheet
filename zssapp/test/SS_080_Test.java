import org.zkoss.ztl.JQuery;


public class SS_080_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        click(jq("$alignRightBtn"));
        waitResponse();
        
        cell_B_8 = getSpecifiedCell(1, 7);
        String textAlign = cell_B_8.css("text-align");
        
        if (textAlign != null) {
            verifyTrue("Unexcepted result: " + textAlign, textAlign.equalsIgnoreCase("right"));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
    }

}
