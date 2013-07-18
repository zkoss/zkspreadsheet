import org.zkoss.ztl.JQuery;


public class SS_063_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        click(jq("$italicBtn"));
        
        cell_B_8 = getSpecifiedCell(1, 7);
        String style = cell_B_8.css("font-style");
        
        if (style != null) {
            verifyTrue("Unexcepted result: " + style, "italic".equalsIgnoreCase(style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
    }

}
