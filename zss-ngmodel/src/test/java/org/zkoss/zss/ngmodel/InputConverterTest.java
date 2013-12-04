package org.zkoss.zss.ngmodel;

import java.text.DateFormat;
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

import junit.framework.Assert;

import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.zss.ngmodel.NCell.CellType;
import org.zkoss.zss.ngmodel.impl.sys.InputConverter;
import org.zkoss.zss.ngmodel.sys.input.InputParseContext;
import org.zkoss.zss.ngmodel.sys.input.InputResult;

public class InputConverterTest {

	static private InputParseContext inputParseContext;
	
	@BeforeClass
	public static void init(){
		inputParseContext = new InputParseContext(Locale.US);
	}
	
	@Test
	public void dateTest(){
		String editText = "01/01/2013";
		DateFormat format = new SimpleDateFormat("dd/MM/yyyy");
		InputResult result = InputConverter.editTextToValue(editText, null, inputParseContext);
		Assert.assertEquals(CellType.NUMBER,result.getType());
		Assert.assertTrue(result.getValue() instanceof Date);
		Assert.assertEquals(editText, format.format(result.getValue()));
	}

	@Test
	public void formulaTest(){
		String editText = "=SUM(10)";
		InputResult result = InputConverter.editTextToValue(editText, null, inputParseContext);
		Assert.assertEquals(CellType.FORMULA,result.getType());
		Assert.assertTrue(result.getValue() instanceof String);
		Assert.assertEquals(editText.substring(1), result.getValue().toString());
	}
}
