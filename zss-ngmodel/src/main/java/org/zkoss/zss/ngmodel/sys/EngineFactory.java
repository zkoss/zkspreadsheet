package org.zkoss.zss.ngmodel.sys;

import org.zkoss.zss.ngmodel.impl.sys.FormulaEngineImpl;
import org.zkoss.zss.ngmodel.impl.sys.InputEngineImpl;
import org.zkoss.zss.ngmodel.sys.formula.FormulaEngine;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;

public class EngineFactory {

	static private EngineFactory instance;

	private EngineFactory() {
	}

	public static EngineFactory getInstance() {
		if (instance == null) {
			synchronized (EngineFactory.class) {
				if (instance == null) {
					// TODO from config
					instance = new EngineFactory();
				}
			}
		}
		return instance;
	}

	public InputEngine getInputEngine() {
		return new InputEngineImpl();
	}

	public FormulaEngine getFormulaEngine() {
		return new FormulaEngineImpl();
	}

}
