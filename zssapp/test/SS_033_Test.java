import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.util.ColorVerifingHelper;

//Format>> Background color

public class SS_033_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        JQuery cell_B_8 = getSpecifiedCell(1, 7);
        clickCell(cell_B_8);
        clickCell(cell_B_8);

    	click("jq('$formatMenu button.z-menu-btn')");
    	waitResponse();
    	click(jq("$backgroundColorMenu a.z-menu-cnt-img"));
    	waitResponse();

    	// Input color hex code, then press Enter.
        JQuery colorTextbox = jq(".z-colorpalette-hex-inp:visible");
        String backgroundColorStr = "#990033";
        type(colorTextbox, backgroundColorStr);
        keyPressEnter(colorTextbox);
    	waitResponse();

        //Verify
        cell_B_8 = getSpecifiedCell(1, 7);
        String bgB8 = cell_B_8.parent().css("background-color");
        System.out.println(">>>>bgB8 = "+bgB8);
        String style = ColorVerifingHelper.transform(bgB8);
        
        if (style != null) {
            verifyTrue("Unexcepted result: " + cell_B_8.css("background-color"), ColorVerifingHelper.isEqualColor(backgroundColorStr, style));
        } else {
            verifyTrue("Cannot get style of specified cell!", false);
        }
   }
}
