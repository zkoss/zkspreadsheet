import org.zkoss.ztl.JQuery;

//focus in cell F21,ctrl+x,focus on cell N11, ctrl+v 
public class SS_051_Test extends SSAbstractTestCase {

    @Override
    protected void executeTest() {
        // Click the same cell twice times because focus always on A1 at first
        // time.
        JQuery cell_F_21 = getSpecifiedCell(5, 20);
        String sourceValue = getCellText(cell_F_21);
        clickCell(cell_F_21);
        clickCell(cell_F_21);

        // Ctrl + X
        keyDownNative(CTRL);
        waitResponse();
        keyDownNative(X);
        waitResponse();
        keyUpNative(X);
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
        cell_F_21 = getSpecifiedCell(5, 20);
        String originalCellValue = cell_F_21.text();
        verifyTrue("F21 current value: " + originalCellValue, originalCellValue.isEmpty());
        String targetValue = getCellText(cell_N_11);
        verifyEquals("Cut value=" + sourceValue + ", Pasted value=" + targetValue, sourceValue, targetValue);
    }

}
