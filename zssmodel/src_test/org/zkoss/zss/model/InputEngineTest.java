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
		//separator without decimal
		String[] editTexts = new String[] {"1,234,5678", "123,45678"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(12345678.0, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("#,##0", result.getFormat());
		}
		
		//one parentheses with comma are legal as negative
		editTexts = new String[] {"(123,4567)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(-1234567.0, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("#,##0", result.getFormat());
		}
		
		//separator with decimal
		editTexts = new String[] {"1,234,567.0", "1,234,567.12", "12,4567.123"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals("#,##0.00", result.getFormat());
		}
		
		//separator distance must > 3 digits(not legal)
		editTexts = new String[] {"1,234,57", "12,456,81.9", "123,45E10"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.STRING,result.getType());
			Assert.assertTrue(result.getValue() instanceof String);
			Assert.assertNull(result.getFormat());
		}
		
		//percentage without decimal
		editTexts = new String[] {"%-123", "-123%", "%(123)", "(123)%", "(%123)", "(123%)", "-((123)%)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(editText, CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(-1.23, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("0%", result.getFormat());
		}
		
		//percentage with decimal
		editTexts = new String[] {"%123.", "123.1%", "123.12%", "123.123%"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals("0.00%", result.getFormat());
		}
		
		//two parentheses(not legal)
		editTexts = new String[] {"%((123))", "((123))%", "((%123))", "((123%))", "-(%(123))"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.STRING,result.getType());
			Assert.assertTrue(result.getValue() instanceof String);
			Assert.assertNull(result.getFormat());
		}
		
		//one parentheses are legal as negative
		editTexts = new String[] {"(123)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(-123.0, ((Double)result.getValue()).doubleValue());
			Assert.assertNull(result.getFormat());
		}
		
		//sign and parentheses are legal
		editTexts = new String[] {"-(123)", "-+(123)", "+-(123)", "---(123)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(-123.0, ((Double)result.getValue()).doubleValue());
			Assert.assertNull(result.getFormat());
		}

		//sign and parentheses are legal
		editTexts = new String[] {"+(123)", "++(123)", "--(123)", "-+-(123)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(123.0, ((Double)result.getValue()).doubleValue());
			Assert.assertNull(result.getFormat());
		}

		//sign and multiple parentheses are legal
		editTexts = new String[] {"+((123))", "++(((123)))", "--((((123))))", "-+-(((((123)))))"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(123.0, ((Double)result.getValue()).doubleValue());
			Assert.assertNull(result.getFormat());
		}

		//sign and parentheses and thousand seperators (not legal)
		editTexts = new String[] {"+(1,234)", "++(1,234)", "--(1,234)", "-+-(1,234)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.STRING,result.getType());
			Assert.assertTrue(result.getValue() instanceof String);
			Assert.assertNull(result.getFormat());
		}

		//currency without decimal
		editTexts = new String[] {"$1234567", "$1,234,567", "$123,4567", "$1234567"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertTrue(result.getValue() instanceof Double);
			Assert.assertEquals(1234567.0, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("$#,##0_);[Red]($#,##0)", result.getFormat());
		}
		
		//currency with decimal
		editTexts = new String[] {"$1234567.89", "$1,234,567.89", "$123,4567.89", "$1234567.89"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertEquals(1234567.89, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("$#,##0.00_);[Red]($#,##0.00)", result.getFormat());
		}
		
		//currency with decimal
		editTexts = new String[] {"$(1234567.89)", "-$1234567.89", "($1,234,567.89)", "$(123,4567.89)", "-$123,4567.89", "$-123,4567.89"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertEquals(-1234567.89, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("$#,##0.00_);[Red]($#,##0.00)", result.getFormat());
		}
		
		//scientific
		editTexts = new String[] {"1234567E1", "1,234567E1", "123,4567E1", "1234567E1", "1,234,567E1"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertEquals(12345670.0, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("0.00E+00", result.getFormat());
		}

		//scientific mixed with currency/parenthesis
		editTexts = new String[] {"(1,234567E1)", "-123,4567E1", "-$1234567E1", "($1,234,567E1)", "$(1,234,567E1)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(CellType.NUMBER,result.getType());
			Assert.assertEquals(-12345670.0, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals("0.00E+00", result.getFormat());
		}
    
	    //exponential mixed with percent
		editTexts = new String[] {"(1,234567E1)%", "%(1,234567E1)"};
		for (String editText : editTexts) {
			InputResult result = inputEngine.parseInput(editText, "", inputParseContext);
			Assert.assertEquals(editText, CellType.NUMBER,result.getType());
			Assert.assertEquals(editText, -123456.70, ((Double)result.getValue()).doubleValue());
			Assert.assertEquals(editText, "0.00E+00", result.getFormat());
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
