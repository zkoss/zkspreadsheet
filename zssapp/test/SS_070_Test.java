import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_070_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select cell
        JQuery cell_L_13 = getSpecifiedCell(11, 12);
        clickCell(cell_L_13);
        clickCell(cell_L_13);
        
        // Click Border icon
        JQuery borderIcon = jq("$borderBtn");
        mouseOver(borderIcon);
        waitResponse();
        clickAt(borderIcon, "30,0");
        waitResponse();
        
        // Click top border
        click(jq(".z-menu-item:eq(1)"));
        waitResponse();
        
        // Verify
        verifyTrue(ColorVerifingHelper.isEqualColor("#000000", getSpecifiedCell(11, 11).css("border-bottom-color")));
    }

}
