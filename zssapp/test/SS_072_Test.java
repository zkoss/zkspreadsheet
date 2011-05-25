import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_072_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select cell
        JQuery cell_L_13 = loadTargetCell();
        clickCell(cell_L_13);
        clickCell(cell_L_13);
        
        // Click Border icon
        JQuery borderIcon = jq("$fastIconBtn $borderBtn:visible");
        mouseOver(borderIcon);
        waitResponse();
        clickAt(borderIcon, "30,0");
        waitResponse();
        
        // Click right border
        click(jq(".z-menu-popup:visible .z-menu-item:eq(3)"));
        waitResponse();
        
        // Verify
        cell_L_13 = loadTargetCell();
        verifyTrue(ColorVerifingHelper.isEqualColor("#000000", cell_L_13.parent().css("border-right-color")));
    }
    
    private JQuery loadTargetCell() {
        return getSpecifiedCell(11, 12);
    }

}
