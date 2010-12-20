import org.zkoss.ztl.JQuery;


public class SS_169_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        click(jq("$fileMenu button.z-menu-btn"));
        waitResponse();
        click(jq("$openFile a.z-menu-item-cnt"));
        waitResponse();
        
        JQuery fileItem = jq("$filesListbox .z-listitem");
        
        if (fileItem.exists()) {
            doubleClick(fileItem);
            waitResponse();
        } else {
            verifyTrue("You must open an excel file before running this case!", false);
        }
        
        sleep(20000);
    }

}
