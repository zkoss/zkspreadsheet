import org.zkoss.ztl.JQuery;


public class SS_050_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Click the same cell twice times because focus always on A1 at first time.
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        clickCell(cell_F_21);
        clickCell(cell_F_21);
        
        // Ctrl + C
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(C);
        waitResponse();
        keyUpNative(C);
        waitResponse();
        keyUpNative(CTRL);
        waitResponse();
        
        // Select another cell
        JQuery cell_N_11 = getSpecifiedCell(13, 10);
        clickCell(cell_N_11);
        
        // Ctrl + V
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(V);
        waitResponse();
        keyUpNative(V);
        waitResponse();
        keyUpNative(CTRL);
        waitResponse();
        
        // Verify value        
        cell_N_11 = getSpecifiedCell(13, 10); // Here must get from dom again.
        String sourceValue = getCellContent(cell_F_21);
        String targetValue = getCellContent(cell_N_11);
        verifyEquals("Copied value=" + sourceValue + ", Pasted value=" + targetValue, sourceValue, targetValue);
    }

}
