package org.zkoss.zss.model;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.util.Locales;
import org.zkoss.zss.model.SCell.CellType;
import org.zkoss.zss.model.sys.EngineFactory;
import org.zkoss.zss.model.sys.input.InputEngine;
import org.zkoss.zss.model.sys.input.InputParseContext;
import org.zkoss.zss.model.sys.input.InputResult;

public class InputEngineTest {

	static private InputEngine inputEngine;
	static private InputParseContext inputParseContext;
	
	@BeforeClass
	public static void init(){
		Setup.touch();
		inputEngine = EngineFactory.getInstance().createInputEngine();
		inputParseContext = new InputParseContext(Locale.US);
	}
	
	@Before
	public void beforeTest() {
		Locales.setThreadLocal(Locale.TAIWAN);
	}
	
	@Test
	public void text(){
		String editText = "=SUM(10)";
		InputResult result = inputEngine.parseInput(editText, "@", inputParseContext);
		Assert.assertEquals(CellType.STRING,result.getType());
		Assert.assertTrue(result.getValue() instanceof String);
		Assert.assertEquals(editText, result.getValue().toString());
	}

	

	@Test
	public void formulaTest(){
		String editText = "=SUM(10)";
		InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
		Assert.assertEquals(CellType.FORMULA,result.getType());
		Assert.assertTrue(result.getValue() instanceof String);
		Assert.assertEquals(editText.substring(1), result.getValue().toString());
	}
	
	
	@Test
	public void booleanType(){
		String editText = "true";
		InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
		Assert.assertEquals(CellType.BOOLEAN,result.getType());
		Assert.assertTrue(result.getValue() instanceof Boolean);
	}
	
	@Test
	public void errorType(){
		String editText = "#NAME?";
		InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
		Assert.assertEquals(CellType.ERROR,result.getType());
		Assert.assertTrue(result.getValue() instanceof Byte);
	}

	@Test
	public void numericType(){
		final Object OK[][] = {
				{"123456", 123456, null},
				{"-123456", -123456, null},
				{"1E+10", 1E+10, null},
				{"-1E+10", -1E+10, null},
				{"1.23456", 1.23456, null},
				{"-1.23456", -1.23456, null},
				{"1E-25", 1E-25, null},
				{"-1E-25", -1E-25, null},
				{"0.025689", 0.025689, null},
				{"-0.025689", -0.025689, null},
		};
		//legal case
		for (int j = 0;  j < OK.length; ++j) {
			Object[] item = OK[j];
			InputResult result = inputEngine.parseInput((String)item[0],  "", inputParseContext);
			Assert.assertEquals((String) item[0] + "->" + j, CellType.NUMBER,result.getType());
			Assert.assertTrue((String) item[0] + "->" + j, result.getValue() instanceof Double);
			Assert.assertEquals((String) item[0] + "->" + j, ((Number) item[1]).doubleValue(), ((Double)result.getValue()).doubleValue(), 0.000000000000001);
			Assert.assertEquals((String) item[0] + "->" + j, (String) item[2], result.getFormat());
		}
		
		final String FAIL[] = {
			"a123", 
			"e123", 
			"12 23",
		};
		//illegal case (a String)
		for (int j = 0; j < FAIL.length; ++j) {
			String item = FAIL[j];
			InputResult result = inputEngine.parseInput(item, "", inputParseContext);
			Assert.assertEquals(item + "->" + j, CellType.STRING,result.getType());
			Assert.assertTrue(item + "->" + j, result.getValue() instanceof String);
			Assert.assertNull(item + "->" + j, result.getFormat());
		}
	}
	
	@Test
	public void dateTest(){
		String editText = "01/01/2013";
		DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
		InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
		Assert.assertEquals(CellType.NUMBER,result.getType());
		Assert.assertTrue(result.getValue() instanceof Date);
		Assert.assertEquals(editText, format.format(result.getValue()));
	}
	
}
