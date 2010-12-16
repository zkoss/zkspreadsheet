import org.zkoss.ztl.JQuery;


public class SS_062_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        click(jq("$boldBtn"));
        
        cell_B_8 = getSpecifiedCell(1, 7);
        String style = cell_B_8.css("font-weight");
        
        if (style != null) {
            if (style.matches("[a-zA-Z]*?")) {
                verifyTrue("Unexcepted result: " + style, "bold".equalsIgnoreCase(style));
            } else if (style.matches("[0-9]*?")) {
                int styleNum = Integer.parseInt(style);
                verifyTrue("Unexcepted result: " + style, styleNum > 400);
            } else {
                verifyTrue("Unknown result: " + style, false);
            }
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
    }

}
