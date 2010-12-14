import org.junit.Test;
import org.zkoss.ztl.JQuery;
import org.zkoss.ztl.ZKClientTestCase;

import com.thoughtworks.selenium.Selenium;

/**
 * An abstract test case class for ZSS. You can just write your test action in your case which
 * extends this class.
 * @author Phoenix
 *
 */
public abstract class SSAbstractTestCase extends ZKClientTestCase {
    public SSAbstractTestCase() {
        target = Utils.getTarget();
        browsers = getBrowsers(Utils.getBrowsers());
        _timeout = Utils.getTimeout();
    }
    
    @Test(expected = AssertionError.class)
    public final void testCase() {
        for (Selenium browser : browsers) {
            try {
                start(browser);
                windowFocus();
                windowMaximize();
                executeTest();
            } finally {
                stop();
            }
        }
    }
    
    /**
     * 
     * @param col - Base on 0
     * @param row - Base on 0
     * @return A JQuery object of cell.
     */
    public JQuery getSpecifiedCell(int col, int row) {
        return jq("div.zscell[z\\\\.c=\"" + col + "\"][z\\\\.r=\"" + row + "\"] div");
    }
    
    /**
     * 
     * @param col - Base on 0
     * @return A JQuery object of column header.
     */
    public JQuery getColumnHeader(int col) {
        return jq("div.zstopcell[z\\\\.c=\"" + col + "\"]div");
    }
    
    /**
     * 
     * @param row - Base on 0
     * @return A JQuery object of row header.
     */
    public JQuery getRowHeader(int row) {
       return jq("div.zsleftcell[z\\\\.r=\"" + row + "\"] div"); 
    }
    
    public String getCellContent(JQuery cellLocator) {
        return cellLocator.text();
    }
    
    public void clickCell(JQuery cellLocator) {
        mouseDownAt(cellLocator, "1,2");
        waitResponse();
        mouseUpAt(cellLocator, "1,2");
        waitResponse();
    }
    
    /**
     * Actually executing test.<br />
     * Note: You also have to implement your validation in this method.
     */
    protected abstract void executeTest();
}
