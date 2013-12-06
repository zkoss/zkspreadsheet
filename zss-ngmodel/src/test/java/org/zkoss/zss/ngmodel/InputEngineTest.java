package org.zkoss.zss.ngmodel;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.BeforeClass;
import org.junit.Ignore;
import org.junit.Test;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.sys.EngineFactory;
import org.zkoss.zss.ngmodel.sys.input.InputEngine;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;

public class InputEngineTest {

	static private InputEngine inputEngine;
	static private InputParseContext inputParseContext;
	
	@BeforeClass
	public static void init(){
		inputEngine = EngineFactory.getInstance().createInputEngine();
		inputParseContext = new InputParseContext(Locale.US);
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
	
	@Ignore("not implemented")
	public void numericTytpe(){
		String editText = "1,234,567";
		InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
		Assert.assertEquals(CellType.NUMBER,result.getType());
		Assert.assertTrue(result.getValue() instanceof Double);
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
