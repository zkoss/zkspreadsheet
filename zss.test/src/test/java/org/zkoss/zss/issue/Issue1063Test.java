package org.zkoss.zss.issue;

import java.io.IOException;
import java.util.List;
import java.util.Locale;

import org.junit.After;
import org.junit.Assert;
import org.junit.Before;
import org.junit.BeforeClass;
import org.junit.Test;
import org.zkoss.poi.ss.util.SheetUtil;
import org.zkoss.zss.Setup;
import org.zkoss.zss.Util;
import org.zkoss.zss.api.Range;
import org.zkoss.zss.api.Ranges;
import org.zkoss.zss.api.model.Book;
import org.zkoss.zss.api.model.Sheet;
import org.zkoss.zss.api.model.Validation;

public class Issue1063Test {
	
	@BeforeClass
	public static void setUpLibrary() throws Exception {
		Setup.touch();
	}
	
	@Before
	public void startUp() throws Exception {
		Setup.pushZssLocale(Locale.TAIWAN);
	}
	
	@After
	public void tearDown() throws Exception {
		Setup.popZssLocale();
	}

	/**
	 * test password encryptor
	 */
	@Test
	public void passwordEncryptor() throws IOException {
		int spinCount=100000;
		String algName = "SHA-512";
		String password = "peter";
		String saltValue = "V6fe+5wecv4/FCzVQYJhmw==";
		String hashValue = "0KY8HEg3mN/Nk5VOi94NWE4jQeV+mOFtOKsvU8/qQr6foNMqEhuW9jQfEh6/H3CyM8t53JOdjInn7u+wKK6h1w==";
		String hashCalc = SheetUtil.encryptPassword(password, algName, saltValue, spinCount);
		
		Assert.assertEquals("hashValue == calculated hash", hashValue, hashCalc);
	}
}