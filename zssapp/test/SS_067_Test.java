import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_067_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select cell
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);
        
        // Click font color button on the toolbar
        click(jq("$cellColorBtn"));
        waitResponse();
        
        // Input color hex code, then press Enter.
        JQuery colorTextbox = jq(".z-colorpalette-hex-inp:eq(1)");
        String bgColorStr = "#990033";
        type(colorTextbox, bgColorStr);
        keyPressEnter(colorTextbox);
        
        //Verify
        cell_B_8 = getSpecifiedCellOuter(1, 7);
        String style = cell_B_8.css("background-color");
        
        if (style != null) {
            verifyTrue("Unexcepted result: " + cell_B_8.css("background-color"), ColorVerifingHelper.isEqualColor(bgColorStr, style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
    }

}
