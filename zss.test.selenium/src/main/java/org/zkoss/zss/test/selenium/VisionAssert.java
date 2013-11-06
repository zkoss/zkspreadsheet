package org.zkoss.zss.test.selenium;

import java.awt.image.BufferedImage;
import java.io.File;
import java.io.FileFilter;
import java.io.IOException;
import java.util.LinkedList;
import java.util.List;

import org.openqa.selenium.OutputType;
import org.openqa.selenium.TakesScreenshot;
import org.openqa.selenium.WebDriver;

public class VisionAssert {

	final String baseName;
	final String resultPrefix;
	final WebDriver driver;

	private List<File> screenshots = new LinkedList<File>();
	// private List<File> results = new LinkedList<File>();
	private List<FailAssertion> assertions = new LinkedList<FailAssertion>();

	public VisionAssert(WebDriver driver, String baseName) {
		this.baseName = baseName;
		this.driver = driver;
		resultPrefix = baseName + "-";
		cleanResultFiles();
	}

	public List<File> getScreenshots() {
		return screenshots;
	}

	// public List<File> getResults(){
	// return results;
	// }

	public List<FailAssertion> getFailAssertioins() {
		return assertions;
	}

	public void cleanScreenshots() {
		screenshots.clear();
	}

	// public void cleanResults(){
	// results.clear();
	// }
	public void cleanFailAssertions() {
		assertions.clear();
	}

	private void cleanResultFiles() {
		File resultFolder = Setup.getResultFolder();
		File[] files = resultFolder.listFiles(new FileFilter() {
			public boolean accept(File pathname) {
				if (pathname.isFile()
						&& pathname.getName().startsWith(resultPrefix)) {
					return true;
				}
				return false;
			}
		});
		for (File f : files) {
			f.delete();
		}
	}

	public void captureOrAssert(String caseName) {
		File screenshotFolder = Setup.getScreenshotFolder();
		File caseScreenshotFile = new File(screenshotFolder, resultPrefix
				+ caseName + ".png");
		if (!caseScreenshotFile.exists()) {
			capture(caseName);
		} else {
			assertIt(caseName);
		}
	}
	
	public void capture(String caseName) {
		File resultFolder = Setup.getResultFolder();
		File screenshotFolder = Setup.getScreenshotFolder();
		File caseScreenshotFile = new File(screenshotFolder, resultPrefix
				+ caseName + ".png");

		File currentScreenshot = ((TakesScreenshot) driver)
				.getScreenshotAs(OutputType.FILE);
		try {
			Files.copy(currentScreenshot, caseScreenshotFile);
			screenshots.add(caseScreenshotFile);
		} catch (IOException e) {
			throw new RuntimeException(e.getMessage(), e);
		}
	}
	
	public void assertIt(String caseName) {
		File resultFolder = Setup.getResultFolder();
		File screenshotFolder = Setup.getScreenshotFolder();
		File caseScreenshotFile = new File(screenshotFolder, resultPrefix
				+ caseName + ".png");
		File caseResultFile = new File(resultFolder, resultPrefix + caseName
				+ ".png");

		File currentScreenshot = ((TakesScreenshot) driver)
				.getScreenshotAs(OutputType.FILE);
		if (!caseScreenshotFile.exists()) {
			assertions.add(
					new FailAssertion("no screenshot to assert",
					null));
		} else {
			// verify
			boolean r = checkScreenshot(caseScreenshotFile, currentScreenshot,
					caseResultFile);
			if (!r) {
				// Assert.fail("screenshot assert fail and store at "+caseResultFile.getAbsolutePath());
				assertions.add(
						new FailAssertion("screenshot assert fail and store at "+ caseResultFile.getAbsolutePath(),
						caseResultFile));
			}
		}
	}

	public static boolean checkScreenshot(File expect, File result,
			File assertFail) {
		Comparator compar = Setup.getImageComparator();
		try {
			BufferedImage img1 = Files.readImage(expect);
			BufferedImage img2 = Files.readImage(result);
			BufferedImage r = compar.compare(img1, img2);
			if (r != null) {
				Files.writeImage(assertFail, r);
				return false;
			}
			return true;
		} catch (Exception x) {
			throw new RuntimeException(x.getMessage(), x);
		}
	}

	static class FailAssertion {
		String message;
		File file;

		public FailAssertion(String message, File file) {
			this.message = message;
			this.file = file;
		}

		public File getFile() {
			return file;
		}

		public void setFile(File file) {
			this.file = file;
		}

		public String getMessage() {
			return message;
		}

		public void setMessage(String message) {
			this.message = message;
		}
	}
}
