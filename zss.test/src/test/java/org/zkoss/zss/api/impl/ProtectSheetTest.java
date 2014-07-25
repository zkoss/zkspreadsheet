package org.zkoss.zss.api.impl;

import java.io.IOException;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.impl.ua.*;
import org.zkoss.zss.ui.impl.undo.HideHeaderAction;
import org.zkoss.zss.ui.sys.SortHandler;
import org.zkoss.zssex.ui.dialog.HeaderSizeCtrl;
import org.zkoss.zssex.ui.impl.ua.*;

/**
 * test Sheet Protection API
 * validation: 
 * Each action handlers should be enable or not when the option is allowed.
 * @author RaymondChao
 *
 */
public class ProtectSheetTest {
	
    private FontFamilyHandler fontFamilyHandler = new FontFamilyHandler();
    private FontSizeHandler fontSizeHandler = new FontSizeHandler();
    private FontBoldHandler fontBoldHandler = new FontBoldHandler();
    private FontItalicHandler fontItalicHandler = new FontItalicHandler();
    private FontUnderlineHandler fontUnderlineHandler = new FontUnderlineHandler();
    private FontStrikeoutHandler fontStrikeoutHandler = new FontStrikeoutHandler();
    private ApplyBorderHandler applyBorderHandler = new ApplyBorderHandler(null, null);
    private FillColorHandler fillColorHandler = new FillColorHandler();
    private FontColorHandler fontColorHandler = new FontColorHandler();
    private VerticalAlignHandler verticalAlignHandler = new VerticalAlignHandler(null);
    private HorizontalAlignHandler horizontalAlignHandler = new HorizontalAlignHandler(null);
    private WrapTextHandler wrapTextHandler = new WrapTextHandler();
    private FormatCellHandler formatCellHandler = new FormatCellHandler();
    private MergeCenterHandler mergeCenterHandler = new MergeCenterHandler();
    private MergeHandler mergeHandler = new MergeHandler(true);
    private UnmergeHandler unmergeHandler = new UnmergeHandler();
    private HeaderSizeHandler columnSizeHandler = new HeaderSizeHandler(HeaderSizeCtrl.HeaderType.COLUMN);
    private HeaderSizeHandler rowSizeHandler = new HeaderSizeHandler(HeaderSizeCtrl.HeaderType.ROW);
    private InsertCellRightHandler insertCellRightHandler = new InsertCellRightHandler();
    private InsertCellDownHandler insertCellDownHandler = new InsertCellDownHandler();
    private ShiftCellHandler shiftCellHandler = new ShiftCellHandler();
    private InsertColumnHandler insertColumnHandler = new InsertColumnHandler();
    private InsertRowHandler insertRowHandler = new InsertRowHandler();
    private InsertFunctionHandler insertFunctionHandler = new InsertFunctionHandler();
    private HyperlinkHandler hyperlinkHandler = new HyperlinkHandler();
    private DeleteCellLeftHandler deleteCellLeftHandler = new DeleteCellLeftHandler();
    private DeleteCellUpHandler deleteCellUpHandler = new DeleteCellUpHandler();
    private DeleteColumnHandler deleteColumnHandler = new DeleteColumnHandler();
    private DeleteRowHandler deleteRowHandler = new DeleteRowHandler();
    private HideHeaderHandler hideColumnHandler = new HideHeaderHandler(HideHeaderAction.Type.COLUMN, true);
    private HideHeaderHandler hideRowHandler = new HideHeaderHandler(HideHeaderAction.Type.ROW, true);
    private ClearCellHandler clearCellHandler = new ClearCellHandler(null);
    private SortHandler sortHandler = new SortHandler(true);
    private CustomSortHandler customSortHandler = new CustomSortHandler();
    private ApplyFilterHandler applyFilterHandler = new ApplyFilterHandler(); 
    private AutoFillHandler autoFillHandler = new AutoFillHandler();
    private CleanFilterHandler cleanFilterHandler = new CleanFilterHandler();
    private ReapplyFilterHandler reapplyFilterHandler = new ReapplyFilterHandler();
    private InsertPictureHandler insertPictureHandler = new InsertPictureHandler();
    private MovePictureHandler movePictureHandler = new MovePictureHandler();
    private DeletePictureHandler deletePictureHandler = new DeletePictureHandler();
    private InsertChartHandler insertChartHandler = new InsertChartHandler(null, null, null);
    private MoveChartHandler moveChartHandler = new MoveChartHandler();
    private DeleteChartHandler deleteChartHandler = new DeleteChartHandler();
    
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}
	
	@Test
	public void testActionHandler() throws IOException {
		final Book book = Util.loadBook("576-protect-sheet.xlsx");
		testProtectAll(book, book.getSheet("sheet1"));
		testAllowAll(book, book.getSheet("sheet2"));
		testProtectAll(book, book.getSheet("sheet3"));
		testAllowFormatCells(book, book.getSheet("sheet4"));
		testAllowFormatColumns(book, book.getSheet("sheet5"));
		testAllowFormatRows(book, book.getSheet("sheet6"));
		testAllowInsertColumns(book, book.getSheet("sheet7"));
		testAllowInsertRows(book, book.getSheet("sheet8"));
		testAllowInsertHyperlinks(book, book.getSheet("sheet9"));
		testAllowDeleteColumns(book, book.getSheet("sheet10"));
		testAllowDeleteRows(book, book.getSheet("sheet11"));
		testSort(book, book.getSheet("sheet12"));
		// TODO PivotTable reports test
		// testAllowPivotTable(book, sheet);
		testEditObjects(book, book.getSheet("sheet13"));
		// TODO edit scenarios test
		// testEditScenarios(book, sheet);
		
		testNull(book, null);
		testNull(null, null);
	}
	
	private void testAllowAll(Book book, Sheet sheet) {
		testAbstracthandler(book, sheet);
		testProtectedhandler(book, sheet);
		testFormatCellsHandler(book, sheet, true);
        testFormatColumnsHandler(book, sheet, true);
        testFormatRowsHandler(book, sheet, true);
        testInsertColumnsHandler(book, sheet, true);
        testInsertRowsHandler(book, sheet, true);
        testHyperlinksHandler(book, sheet, true);
        testDeleteColumnsHandler(book, sheet, true);
        testDeleteRowsHandler(book, sheet, true);
        testSortHandler(book, sheet, true);
        testObjectsHandler(book, sheet, true);
	}
	
	private void testProtectAll(Book book, Sheet sheet) {
		testAbstracthandler(book, sheet);
		testProtectedhandler(book, sheet);
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);
	}
	
	private void testAllowFormatCells(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, true);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);
	}
	
	private void testAllowFormatColumns(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, true);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);
	}
	
	private void testAllowFormatRows(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, true);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);

	}
	
	private void testAllowInsertColumns(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, true);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);
	}
	
	private void testAllowInsertRows(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, true);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);

	}
	
	private void testAllowInsertHyperlinks(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, true);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);

	}
	
	private void testAllowDeleteColumns(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, true);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);
	}
	
	private void testAllowDeleteRows(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, true);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);
	}
	
	private void testSort(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, true);
        testObjectsHandler(book, sheet, false);
	}
	
	private void testEditObjects(Book book, Sheet sheet) {
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, true);
	}
	
	private void testFormatCellsHandler(Book book, Sheet sheet, boolean enabled) {
		Assert.assertEquals(fontFamilyHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(fontSizeHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(fontBoldHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(fontItalicHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(fontUnderlineHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(fontStrikeoutHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(applyBorderHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(fillColorHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(fontColorHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(verticalAlignHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(horizontalAlignHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(wrapTextHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(autoFillHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(formatCellHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testFormatColumnsHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(columnSizeHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(hideColumnHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testFormatRowsHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(rowSizeHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(hideRowHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testInsertColumnsHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(insertColumnHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testInsertRowsHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(insertRowHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testHyperlinksHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(hyperlinkHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testDeleteColumnsHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(deleteColumnHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testDeleteRowsHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(deleteRowHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testSortHandler(Book book, Sheet sheet, boolean enabled) {
        Assert.assertEquals(sortHandler.isEnabled(book, sheet), enabled);
        Assert.assertEquals(customSortHandler.isEnabled(book, sheet), enabled);
	}

	private void testObjectsHandler(Book book, Sheet sheet, boolean enabled) {
		Assert.assertEquals(insertPictureHandler.isEnabled(book, sheet), enabled);
		Assert.assertEquals(movePictureHandler.isEnabled(book, sheet), enabled);
		Assert.assertEquals(deletePictureHandler.isEnabled(book, sheet), enabled);
		Assert.assertEquals(insertChartHandler.isEnabled(book, sheet), enabled);
		Assert.assertEquals(moveChartHandler.isEnabled(book, sheet), enabled);
		Assert.assertEquals(deleteChartHandler.isEnabled(book, sheet), enabled);
	}
	
	private void testAbstracthandler(Book book, Sheet sheet) {
		Assert.assertTrue(clearCellHandler.isEnabled(book, sheet));
	}
	
	private void testProtectedhandler(Book book, Sheet sheet) {
		Assert.assertFalse(mergeCenterHandler.isEnabled(book, sheet));
		Assert.assertFalse(mergeHandler.isEnabled(book, sheet));
		Assert.assertFalse(unmergeHandler.isEnabled(book, sheet));
        Assert.assertFalse(insertCellRightHandler.isEnabled(book, sheet));
        Assert.assertFalse(insertCellDownHandler.isEnabled(book, sheet));
        Assert.assertFalse(shiftCellHandler.isEnabled(book, sheet));
        Assert.assertFalse(insertFunctionHandler.isEnabled(book, sheet));
        Assert.assertFalse(deleteCellLeftHandler.isEnabled(book, sheet));
        Assert.assertFalse(deleteCellUpHandler.isEnabled(book, sheet));
        Assert.assertFalse(applyFilterHandler.isEnabled(book, sheet));
        Assert.assertFalse(cleanFilterHandler.isEnabled(book, sheet));
        Assert.assertFalse(reapplyFilterHandler.isEnabled(book, sheet));
	}
	
	private void testNull(Book book, Sheet sheet) {
		testProtectedhandler(book, sheet);
		testFormatCellsHandler(book, sheet, false);
        testFormatColumnsHandler(book, sheet, false);
        testFormatRowsHandler(book, sheet, false);
        testInsertColumnsHandler(book, sheet, false);
        testInsertRowsHandler(book, sheet, false);
        testHyperlinksHandler(book, sheet, false);
        testDeleteColumnsHandler(book, sheet, false);
        testDeleteRowsHandler(book, sheet, false);
        testSortHandler(book, sheet, false);
        testObjectsHandler(book, sheet, false);
	}
}
