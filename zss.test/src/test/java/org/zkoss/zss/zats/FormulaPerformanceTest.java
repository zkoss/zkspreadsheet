package org.zkoss.zss.zats;

import org.junit.*;
import org.junit.internal.AssumptionViolatedException;
import org.junit.rules.Stopwatch;
import org.junit.runner.Description;
import org.zkoss.zats.mimic.*;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.*;
import org.zkoss.zss.ui.Spreadsheet;

import java.util.concurrent.TimeUnit;

import static org.junit.Assert.assertEquals;
import static org.junit.Assert.assertTrue;


/**
 * @author Hawk
 *
 */
public class FormulaPerformanceTest extends SpreadsheetTestCaseBase{

	@Test
    public void importFile(){
		Zats.newClient().connect("/performance/formula.zul");
		assertTrue(stopwatch.runtime(TimeUnit.MILLISECONDS) < 12000 );
    }


	@Rule
	public final Stopwatch stopwatch = new Stopwatch() {
		protected void succeeded(long nanos, Description description) {
			System.out.println(description.getMethodName() + " succeeded, time taken " + toMillis(nanos));
		}

		/**
		 * Invoked when a test fails
		 */
		protected void failed(long nanos, Throwable e, Description description) {
			System.out.println(description.getMethodName() + " failed, time taken " + toMillis(nanos));
		}

		/**
		 * Invoked when a test is skipped due to a failed assumption.
		 */
		protected void skipped(long nanos, AssumptionViolatedException e,
							   Description description) {
			System.out.println(description.getMethodName() + " skipped, time taken " + toMillis(nanos));
		}

		/**
		 * Invoked when a test method finishes (whether passing or failing)
		 */
		protected void finished(long nanos, Description description) {
			System.out.println(description.getMethodName() + " finished, time taken " + toMillis(nanos));
		}

		private String toMillis(long nanos){
			return TimeUnit.MILLISECONDS.convert(nanos, TimeUnit.NANOSECONDS) + " " + TimeUnit.MILLISECONDS.toString();
		}
	};
	
}
