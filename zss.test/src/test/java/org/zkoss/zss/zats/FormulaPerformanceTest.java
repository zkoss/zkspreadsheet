package org.zkoss.zss.zats;

import org.junit.*;
import org.junit.rules.Stopwatch;
import org.zkoss.zats.mimic.*;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.ui.Spreadsheet;

import java.util.concurrent.TimeUnit;

import static org.junit.Assert.assertTrue;


/**
 * @author Hawk
 *
 */
public class FormulaPerformanceTest extends SpreadsheetTestCaseBase{

	private Spreadsheet spreadsheet;
	private DesktopAgent desktop;
	private ComponentAgent zss;

	@Test
    public void importFile(){
		Zats.newClient().connect("/performance/formula.zul");
		assertTrue(stopwatch.runtime(TimeUnit.MILLISECONDS) < 12000 );
    }

	@Test
	public void editCells(){
		desktop =  Zats.newClient().connect("/performance/formula.zul");
		zss = desktop.query("spreadsheet");
		spreadsheet = zss.as(Spreadsheet.class);
		assertTrue(spreadsheet.getSelectedSheetName().equals("录入表"));
		int startingRow = 4;
		int nCells = 1;
		Sheet sheet = spreadsheet.getSelectedSheet();
		long startTime =  stopwatch.runtime(TimeUnit.MILLISECONDS);
		for (int row = startingRow ; row < startingRow + nCells ; row++ ) {
			Ranges.range(sheet, row, 7).setCellEditText("1");
		}
		long endTime =  stopwatch.runtime(TimeUnit.MILLISECONDS);
		assertTrue(endTime - startTime < 80000 );
	}    


	@Rule
	public final Stopwatch stopwatch = new PrintStopwatch();

}
