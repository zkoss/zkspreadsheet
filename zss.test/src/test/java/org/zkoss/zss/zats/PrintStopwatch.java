package org.zkoss.zss.zats;

import org.junit.internal.AssumptionViolatedException;
import org.junit.rules.Stopwatch;
import org.junit.runner.Description;

import java.util.concurrent.TimeUnit;

public class PrintStopwatch extends Stopwatch {
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
}
