/* BugTestingEnvironment.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Apr 10, 2012 4:43:21 PM , Created by sam
}}IS_NOTE

Copyright (C) 2012 Potix Corporation. All Rights Reserved.

{{IS_RIGHT
}}IS_RIGHT
*/
package org.zkoss.test;

import org.zkoss.test.zss.Cell;
import org.zkoss.test.zss.CellBlock;
import org.zkoss.test.zss.CellCache;
import org.zkoss.test.zss.CellCacheAggeration;

import com.google.inject.assistedinject.FactoryModuleBuilder;

/**
 * @author sam
 *
 */
public class BugTestingEnvironment extends TestingEnvironment {

	@Override
	protected void configure() {
		
		install(new FactoryModuleBuilder().build(Cell.Factory.class));
		install(new FactoryModuleBuilder().build(CellBlock.Factory.class));
		install(new FactoryModuleBuilder().build(CellCache.Factory.class));
		install(new FactoryModuleBuilder().build(CellCacheAggeration.BuilderFactory.class));
		
		super.configure();
	}
	
}
