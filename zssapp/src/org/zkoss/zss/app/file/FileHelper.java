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
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Map;

import org.zkoss.io.Files;
import org.zkoss.lang.Library;
import org.zkoss.util.media.Media;
import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Executions;
import org.zkoss.zk.ui.UiException;
import org.zkoss.zss.app.Consts;
import org.zkoss.zss.model.Book;
import org.zkoss.zss.model.Exporter;
import org.zkoss.zss.model.Exporters;
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
	
	/*has import permission or not, default is false*/
	public final static String KEY_IMPORT_PERMISSION = "org.zkoss.zss.app.file.fileHelper.importPermission";
	
	/*has save permission or not, default is false*/
	public final static String KEY_SAVE_PERMISSION = "org.zkoss.zss.app.file.fileHelper.savePermission";
	
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
	
	public static boolean openSrc(String src, Spreadsheet spreadsheet) {
		String key = removeFolderPath(src);
		Map<String, SpreadSheetMetaInfo> infos = SpreadSheetMetaInfo.getMetaInfos();
		if (infos.containsKey(key)) {
			return openSpreadsheet(spreadsheet, infos.get(key));
		}
		return false;
	}
	
	private static String removeFolderPath(String src) {
		int idx = -1;
		String fileName = src;
		if ((idx = fileName.lastIndexOf("\\")) >= 0 || (idx = fileName.lastIndexOf("/")) >= 0) {
			return fileName.substring(idx + 1);
		}
		return fileName;
	}
	
	public static boolean openSpreadsheet(Spreadsheet ss, SpreadSheetMetaInfo info) {
		FileInputStream input = null;
		try {
			input = new FileInputStream(getSpreadsheetStorageFolderPath() + info.getHashFileName());
			ss.setBookFromStream(input, info.getFileName());
			return true;
		} catch (FileNotFoundException e) {
			try {
				Messagebox.show("Can not find file: " + info.getFileName());
			} catch (InterruptedException e1) {
			}
		} finally {
			if (input != null)
				try {
					input.close();
				} catch (IOException e) {
					throw new RuntimeException(e);
				}
		}
		return false;
	}
	
	public static void openNewSpreadsheet(Spreadsheet ss) {
		FileInputStream input = null;
		try {
			input = new FileInputStream(getSpreadsheetStorageFolderPath() + EMPTY_FILE_NAME);
			ss.setBookFromStream(input, EMPTY_FILE_NAME);
		} catch (FileNotFoundException e) {
			throw new RuntimeException(e);
		} finally {
			if (input != null)
				try {
					input.close();
				} catch (IOException e) {
					throw new RuntimeException(e);
				}
		}
	}
	
	public static void saveSpreadsheet(Spreadsheet spreadsheet) {
		String filename = spreadsheet.getSrc();
		SpreadSheetMetaInfo info = SpreadSheetMetaInfo.newInstance(filename);
		Book wb = spreadsheet.getBook();
		Exporter c = Exporters.getExporter("excel");
		
		String fileName = getSpreadsheetStorageFolderPath() + info.getHashFileName();
		File file = new File(fileName);		
		FileOutputStream out = null;
		try {
			file.createNewFile();
			out = new FileOutputStream(file);
			c.export(wb, out);
			SpreadSheetMetaInfo.add(info);
		} catch (IOException e) {
			try {
				Messagebox.show("Save excel failed");
				e.printStackTrace();
			} catch (InterruptedException e1) {
			}
			return;
		} finally {
			if (out != null)
				try {
					out.close();
				} catch (IOException e) {
					throw new RuntimeException(e);
				}
		}
	}
	
	public static void saveSpreadsheetAs(Spreadsheet ss, String fileName) {
		throw new UiException("save file not implmented yet");

//		SpreadSheetMetaInfo info = SpreadSheetMetaInfo.newInstance(filename);
//		Exporter exporter = 
//		exporter.exports
//		SpreadSheetMetaInfo.add(info);
	}
	
	public static void deleteSpreadsheet(String src) {
		Map<String, SpreadSheetMetaInfo> infos = SpreadSheetMetaInfo.getMetaInfos();
		SpreadSheetMetaInfo info = null;
		if (infos.containsKey(src)) {
			info = infos.get(src);
			deleteSpreadSheet(info);
		}
	}
	
	public static void deleteSpreadsheet(Spreadsheet ss) {
		Map<String, SpreadSheetMetaInfo> infos = SpreadSheetMetaInfo.getMetaInfos();
		SpreadSheetMetaInfo info = null;
		if (infos.containsKey(ss.getSrc())) {
			info = infos.get(ss.getSrc());
			deleteSpreadSheet(info);
		}
	}
	
	public static void deleteSpreadSheet(SpreadSheetMetaInfo info) {
		try {
			SpreadSheetMetaInfo.delete(info);
		} catch (IOException e) {
			try {
				Messagebox.show("Delete file failed");
			} catch (InterruptedException e1) {
			}
		}
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
	 * Returns whether has save permission or not
	 * <p> Default: false
	 * @return
	 */
	public static boolean hasSavePermission() {
		return "true".equals(Library.getProperty(KEY_SAVE_PERMISSION));
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
}