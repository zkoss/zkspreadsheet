package org.zkoss.zss.test.selenium.testcases.function;

import org.junit.*;
import org.openqa.selenium.*;
import org.zkoss.zss.test.selenium.*;
import org.zkoss.zss.test.selenium.entity.*;
import org.zkoss.zss.test.selenium.util.*;

import static org.zkoss.zss.test.selenium.util.MyPngCropper.TOP_BOTTOM_CROPPER;


public class AlignmentOverflowTest extends ZSSTestCase {

    private static final String OVERFLOW_RIGHT_ALIGN_ZUL = "/alignment/1338-overflow-right-align.zul";
    private static final String RIGHT_ALIGN_FROZEN_ZUL = "/alignment/1384-right-align-frozen.zul";
    private static final String ALIGNMENT_ZUL = "/alignment/alignment.zul";

    /*
     * verify the importing result
     */
    @Test
    public void testZSS1338Import() throws Exception {
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        captureOrAssert("import", TOP_BOTTOM_CROPPER);
    }

    // always move selection box to A1 before capturing a screenshot to avoid image comparison being affected by selection box
    @Test
    public void testZSS1338ChangeAlign() throws Exception {
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell("F5"));
        dragAndDrop(ctrl.getCell("F5").toWebElement(), ctrl.getCell("F10").toWebElement());
        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_LEFT.toString());
        focusSheet();
        captureOrAssert("changeToLeftAlign", TOP_BOTTOM_CROPPER);

        dragAndDrop(ctrl.getCell("F5").toWebElement(), ctrl.getCell("F10").toWebElement());
        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_RIGHT.toString());
        focusSheet();
        captureOrAssert("changeToRightAlign", TOP_BOTTOM_CROPPER);
    }

    @Test
    public void testZSS1338ChangeAlignMergedCell() throws Exception {
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget ctrl = ss.getSheetCtrl();
        click(ctrl.getCell("E11"));
        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_LEFT.toString());
        focusSheet();
        captureOrAssert("changeToLeftAlign", TOP_BOTTOM_CROPPER);

        click(ZSStyle.HORIZONTAL_ALIGN + " .zstbtn-cave");
        click(ZSStyle.MENUPOPUP_OPEN.toString() + " " + ZSStyle.ALIGN_RIGHT.toString());
        focusSheet();
        captureOrAssert("changeToRightAlign", TOP_BOTTOM_CROPPER);
    }

    @Test
    public void testZSS1338PartialOverflow() throws Exception {
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        doubleClick(sheetCtrl.getCell("D9"));
        waitForTime(Setup.getTimeoutL0());
        EditorWidget editor = sheetCtrl.getInlineEditor();
        type(editor, "cut");
        captureOrAssert("partial1");
        doubleClick(sheetCtrl.getCell("E9"));
        type(editor, "cut");
        waitForTime(Setup.getTimeoutL0());
        captureOrAssert("partial2");
    }


    @Test
    public void testCutOverflowLeftAlignment() throws Exception {
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        CellWidget e1 = sheetCtrl.getCell("E1");
        doubleClick(e1);
        waitForTime(Setup.getTimeoutL0());
        EditorWidget editor = sheetCtrl.getInlineEditor();
        type(editor, "cut");
        captureOrAssert("cutLeft1");
        click(e1);
        delete(editor);
        captureOrAssert("cutLeft2");
    }

    @Test
    public void testZSS1364SheetSwitching(){
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        sheetFunction().gotoTab(2);
        waitForTime(Setup.getTimeoutL0());
        sheetFunction().gotoTab(1);
        waitForTime(Setup.getTimeoutL0());
        captureOrAssert("1", TOP_BOTTOM_CROPPER);

    }

    @Test
    public void testZSS1364EditSheetSwitching() {
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        SheetCtrlWidget sheetCtrl = ss.getSheetCtrl();
        doubleClick(sheetCtrl.getCell("D5"));
        waitForTime(Setup.getTimeoutL0());
        EditorWidget editor = sheetCtrl.getInlineEditor();
        sendKeys(editor, "cut");
        sendKeys(editor, Keys.chord(Keys.ENTER));
        sheetFunction().gotoTab(2);
        waitForTime(Setup.getTimeoutL0());
        sheetFunction().gotoTab(1);
        waitForTime(Setup.getTimeoutL0());
        captureOrAssert("1", TOP_BOTTOM_CROPPER);
    }

    @Test //ZSS-1364
    public void notRenderedCellAtFirst(){
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        setZSSScrollTop(1400);
        waitForTime(Setup.getTimeoutL1());
        captureOrAssert("1", new MyPngCropper(100, 600, 400, 0));
    }

    @Test //ZSS-1375
    public void resizeColumnProtectedSheet(){
        getTo(OVERFLOW_RIGHT_ALIGN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        sheetFunction().gotoTab(3);
        waitForTime(Setup.getTimeoutL1());
        //enlarge a column width
        dragAndDropBy(jq(".zshbouni").get(5), 100, 0);
        click(ss.getSheetCtrl().getCell("A1"));
        waitForTime(Setup.getTimeoutL1());
        captureOrAssert("1", TOP_BOTTOM_CROPPER);
    }

    @Test //ZSS-1384
    public void mergedAcrossFrozenColumn() {
        getTo(RIGHT_ALIGN_FROZEN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        waitForTime(Setup.getTimeoutL0());
        captureOrAssert("1", TOP_BOTTOM_CROPPER);
    }

    @Test //ZSS-1384
    public void mergedAcrossFrozenColumnSwitchSheets() {
        getTo(RIGHT_ALIGN_FROZEN_ZUL);
        basename();
        SpreadsheetWidget ss = focusSheet();
        sheetFunction().gotoTab(2);
        waitForTime(Setup.getTimeoutL0());
        sheetFunction().gotoTab(1);
        waitForTime(Setup.getTimeoutL0());
        captureOrAssert("1", TOP_BOTTOM_CROPPER);
    }

    @Test
    public void testZSS1388Import() throws Exception {
        getTo(ALIGNMENT_ZUL);
        basename();
        captureOrAssert("import", TOP_BOTTOM_CROPPER);
    }
}





