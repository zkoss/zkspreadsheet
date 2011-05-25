import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;


public class SS_077_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Select cells
        clickCell(loadCellL13());
        clickCell(loadCellL13());
        mouseDownAt(loadCellL13(), "1,2");
        waitResponse();
        mouseMoveAt(loadCellM14(), "1,2");
        mouseUp(loadCellM14());
        waitResponse();
        
        // Click Border icon
        JQuery borderIcon = jq("$fastIconBtn $borderBtn:visible");
        mouseOver(borderIcon);
        waitResponse();
        clickAt(borderIcon, "30,0");
        waitResponse();
        
        // Click inside horizontal border
        click(jq(".z-menu-popup:visible .z-menu-item:eq(8)"));
        waitResponse();
        
        // Verify
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", loadCellL13().parent().css("border-right-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#000000", loadCellL13().parent().css("border-bottom-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", loadCellM13().parent().css("border-right-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#000000", loadCellM13().parent().css("border-bottom-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", loadCellL14().parent().css("border-right-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", loadCellL14().parent().css("border-bottom-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", loadCellM14().parent().css("border-right-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", loadCellM14().parent().css("border-bottom-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", getSpecifiedCell(10, 12).parent().css("border-right-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", getSpecifiedCell(10, 13).parent().css("border-right-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", getSpecifiedCell(11, 11).parent().css("border-bottom-color")));
        verifyTrue(ColorVerifingHelper.isEqualColor("#FFFFFF", getSpecifiedCell(12, 11).parent().css("border-bottom-color")));
    }
    
    private JQuery loadCellL14() {
        return getSpecifiedCell(11, 13);
    }

    private JQuery loadCellM13() {
        return getSpecifiedCell(12, 12);
    }

    private JQuery loadCellL13() {
        return getSpecifiedCell(11, 12);
    }
    
    private JQuery loadCellM14() {
        return getSpecifiedCell(12, 13);
    }

}
