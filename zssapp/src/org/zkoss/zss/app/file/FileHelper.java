/* ImportHelper.java

{{IS_NOTE
	Purpose:
		
	Description:
		
	History:
		Oct 29, 2010 8:10:22 PM , Created by Sam
}}IS_NOTE

Copyright (C) 2009 Potix Corporation. All Rights Reserved.

*/
package org.zkoss.zss.app.file;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.io.InputStream;

import org.zkoss.io.Files;
import org.zkoss.lang.Library;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.app.zul.Zssapps;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Messagebox;

/**
 * A spreadsheet file helper class for import, open, save, save as, delete file.
 * 
 * 
 * @author Sam
 *
 */
public class FileHelper {
	private FileHelper(){}
	
	private final static String[] SUPPORTED_FORMAT = new String[]{"xls", "xlsx"};
	
	/*current opening file name*/
	private final static String KEY_OPENING_FILE = "org.zkoss.zss.app.file.fileHelper.openingFile";
	
	/*whether has import permission or not*/
	public final static String KEY_IMPORT_PERMISSION = "org.zkoss.zss.app.file.fileHelper.importPermission";
	
	/*absolute file path that store all path*/
	private static String storageFolderPath;
	
	private final static String EMPTY_FILE_NAME =  "Untitled";
	/**
	 * Store spreadsheet
	 * @param spreadSheetSource
	 */
	public static SpreadSheetMetaInfo store( Media media)
		throws UnsupportedSpreadSheetFileException {
				
		if (!isSupportedSpreadSheetExtention(media.getName())) {
			throw new UnsupportedSpreadSheetFileException(
					"the file extension is incorrect. media name:"+media.getName());
		}
		
		Media spreadSheetSource = media;
		
		SpreadSheetMetaInfo info = 
			SpreadSheetMetaInfo.newInstance(spreadSheetSource.getName()); 
		
		InputStream inputStream = spreadSheetSource.getStreamData();
		try {
			Files.copy(new File(getSpreadsheetStorageFolderPath() + info.getHashFileName()), 
					inputStream);

			SpreadSheetMetaInfo.add(info);
		} catch (IOException e) {
			info = null;
			throw new RuntimeException(e);
		} finally {
			if (inputStream != null) {
				try {
					inputStream.close();
				} catch (IOException e) {
					throw new RuntimeException(e);
				}
			}
			return info;
		}
	}
	
	public static void openSpreadsheet(Spreadsheet ss, SpreadSheetMetaInfo info) {
		FileInputStream input = null;
		try {
			input = new FileInputStream(getSpreadsheetStorageFolderPath() + info.getHashFileName());
			ss.setBookFromStream(input, info.getFileName());
		} catch (FileNotFoundException e) {
			try {
				Messagebox.show("Can not find file: " + info.getFileName());
			} catch (InterruptedException e1) {
			}
		}
	}
	
	public static void openNewSpreadsheet(Spreadsheet ss) {
		
		FileInputStream input = null;
		try {
			input = new FileInputStream(getSpreadsheetStorageFolderPath() + EMPTY_FILE_NAME);
			ss.setBookFromStream(input, EMPTY_FILE_NAME);
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e);
		}
	}
	
	public static void saveSpreadsheet(Spreadsheet ss) {
		throw new UiException("save file not implmented yet");	

//		String filename = ss.getSrc();
//		
//		SpreadSheetMetaInfo info = SpreadSheetMetaInfo.newInstance(filename);
//
//		ExcelExporter exporter = new ExcelExporter();
//		exporter.exports(ss.getBook(), new File(getSpreadsheetStorageFolderPath() + info.getHashFileName()));
//		SpreadSheetMetaInfo.add(info);
	}
	
	public static void saveSpreadsheetAs(Spreadsheet ss, String fileName) {
		throw new UiException("save file not implmented yet");

//		SpreadSheetMetaInfo info = SpreadSheetMetaInfo.newInstance(filename);
//		Exporter exporter = 
//		exporter.exports
//		SpreadSheetMetaInfo.add(info);
	}
	
	public static void deleteSpreadsheet(Spreadsheet ss) {
		throw new UiException("delete file not implmented yet");
//		Map<String, SpreadSheetMetaInfo> infos = SpreadSheetMetaInfo.getMetaInfos();
//		if (infos.containsKey(ss.getSrc())) {
//			
//		}
	}

	/**
	 * Returns spreadsheet supported file format
	 * @return string[]
	 */
	public static String[] getSupportedFormat() {
		return SUPPORTED_FORMAT;
	}
	/**
	 * Returns whether support this extention file or not
	 * @param string file name
	 * @return
	 */
	private static boolean isSupportedSpreadSheetExtention(String fileName) {
		//Note: <= 0, file name at least has one character
		if (fileName == null || fileName.lastIndexOf(".") <= 0)
			return false;

		String extName = fileName.substring(fileName.lastIndexOf(".") + 1);

		for (int i = 0; i < SUPPORTED_FORMAT.length; i++) {
			if (SUPPORTED_FORMAT[i].equalsIgnoreCase(extName)) {
				return true;
			} else
				continue;
		}
		return false;
	}
	/**
	 * Returns the media's extension name
	 * @param string file name
	 * @return extension name or null if there is no extension name
	 */
	public static String getMediaExtention(String fileName) {
		int dotIdx = fileName.indexOf(".");
		String extName = dotIdx >= 0 ? fileName.substring(dotIdx + 1) : null;
		return extName;
	}
	
	public static File createMetaInfoFileIfNeeded() {
		File file = new File(getSpreadsheetStorageFolderPath() + "metaFile");
		if (!file.exists())
			try {
				file.createNewFile();
				return file;
			} catch (IOException e) {
				throw new RuntimeException(e);
			}
		return file;
	}
	
	/**
	 * TODO: move this method to permission control object
	 * 
	 * Returns whether has import permission or not
	 * <p> Default: false
	 * @return
	 */
	public static boolean hasImportPermission() {
		return "true".equals(Library.getProperty(KEY_IMPORT_PERMISSION));
	}
	
	/**
	 * Returns absolute file path that storage spreadsheet
	 * @return
	 */
	public static String getSpreadsheetStorageFolderPath() {
		if (storageFolderPath == null)
			storageFolderPath = Executions.getCurrent().getDesktop().getWebApp().getRealPath("xls") + File.separator;
			
		return storageFolderPath;
	}
	
	public static void createOpenFileDialog(Component parent) {
		Executions.createComponents(Consts._FileListOpen_zul, parent, null);
	}
	
	public static void createImportFileDialog(Component parent) {
		Executions.createComponents(Consts._ImportFile_zul,	parent, null);
	}
}