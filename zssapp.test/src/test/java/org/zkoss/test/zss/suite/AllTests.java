/* AllTests.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Jan 19, 2012 2:40:21 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test.zss.suite;

import java.util.Set;

import org.junit.runner.RunWith;
import org.junit.runner.notification.Failure;
import org.junit.runner.notification.RunListener;
import org.junit.runner.notification.RunNotifier;
import org.junit.runners.Suite;
import org.reflections.Reflections;
import org.zkoss.test.zss.ZSSTestCase;

/**
 * @author sam
 *
 */
@RunWith(AllTests.AllTestsRunner.class)
public final class AllTests {
  private AllTests() {
    // static only
  }
 
  /**
   * Run zss tests cases
   */
  public static class AllTestsRunner extends Suite {
 
    public AllTestsRunner(final Class<?> clazz) throws org.junit.runners.model.InitializationError {
      super(clazz, findClasses());
    }
 
    @Override
    public void run(final RunNotifier notifier) {
    	notifier.addListener(new RunListener() {
    		@Override
    		public void testFailure(Failure failure) throws Exception {
    			// TODO get logger and log it
    		}

    		@Override
    		public void testAssumptionFailure(Failure failure) {
    			// TODO get logger and log it
    		}
    	});
    	super.run(notifier);
    }
 
    private static Class<?>[] findClasses() {
    	Reflections reflections = new Reflections("org.zkoss.test.zss.cases");
    	Set<Class<?>> testCases = reflections.getTypesAnnotatedWith(ZSSTestCase.class);
    	return testCases.toArray(new Class[testCases.size()]);
    }
  }
}