import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_070_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select cell
        JQuery cell_L_13 = getSpecifiedCell(11, 12);
        clickCell(cell_L_13);
        clickCell(cell_L_13);
        
        focusOnCell(11, 12);
        clickDropdownButtonMenu("$fastIconBtn $borderBtn", "Top border");
        
        // Verify
        verifyTrue(ColorVerifingHelper.isEqualColor("#000000", getSpecifiedCell(11, 11).parent().css("border-bottom-color")));
    }

}
