package org.zkoss.zss.app;

import java.io.BufferedReader;
import java.io.BufferedWriter;
import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.FileReader;
import java.io.FileWriter;
import java.io.IOException;
import java.util.ArrayList;
import java.util.Collection;
import java.util.Date;
import java.util.HashMap;
import java.util.Iterator;
import java.util.LinkedList;
import java.util.List;
import java.util.Map;
import java.util.Properties;
import java.util.Queue;
import java.util.Map.Entry;

import org.zkoss.zk.ui.Component;
import org.zkoss.zk.ui.Path;
import org.zkoss.zk.ui.UiException;
//import org.zkoss.zss.model.impl.ExcelExporter;
import org.zkoss.zss.ui.Spreadsheet;
import org.zkoss.zul.Label;
import org.zkoss.zul.Listbox;
import org.zkoss.zul.Listcell;
import org.zkoss.zul.Listitem;
import org.zkoss.zul.Textbox;
import org.zkoss.zul.Window;

//import com.mysql.jdbc.PreparedStatement;

public class FileHelper {
	static String xlsDir;
	// TODO: will this limit concurrent opening spreadsheet to 1?
	static String openingFile;// filename without folder name
	Spreadsheet spreadsheet;
	

	public FileHelper(Spreadsheet spreadsheet) {
		this.spreadsheet = spreadsheet;
		openingFile = spreadsheet.getSrc();
		xlsDir = spreadsheet.getDesktop().getWebApp().getRealPath("xls") + System.getProperty("file.separator");
	}

	public void selectDispatcher(String type) {
		if (type.equals("open")) {
			onOpenFileSelected();
		} else if (type.equals("save")) {
			onSaveFile();
		} else if (type.equals("saveAs")) {
			onSaveAsFileSelected();
		} else if (type.equals("delete")) {
			onDeleteFileSelected();
		} else if (type.equals("export")) {
			onExportFileSelected();
		}
	}

	public void dispatcher(String type) {
		if (type.equals("open")) {
		//	onOpenFile();
		} else if (type.equals("save")) {
			onSaveFile();
		} else if (type.equals("saveAs")) {
			onSaveAsFile();
		} 
	}
	
	public void onSaveFile() {
//TODO save file
throw new UiException("save file not implmented yet");
/*
		BufferedReader reader = null;
		BufferedWriter writer = null;
		try {
			String filename = spreadsheet.getSrc();
			ExcelExporter exporter = new ExcelExporter();

			long millis = System.currentTimeMillis();
			String hashFilename = millis + ".xls";
			exporter.exports(spreadsheet.getBook(), new File(xlsDir	+ hashFilename));

			Properties prop = new Properties();
			reader = new BufferedReader(new FileReader(xlsDir + "metaFile"));
			prop.load(reader);
			reader.close();
			reader = null;

			prop.put(filename,  millis + ", " + filename + ", "	+ hashFilename);
			writer = new BufferedWriter(new FileWriter(xlsDir + "metaFile", false));
			prop.store(writer, null);
			writer.close();
			writer = null;
			
			MainWindowCtrl w1 = MainWindowCtrl.getInstance();
			w1.hashFilename = hashFilename;
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (reader != null) {
				try {
					reader.close();
				} catch (IOException e) {
				}
			}
			if (writer != null) {
				try {
					writer.close();
				} catch (IOException e) {
				}
			}
		}
*/	}
	
	public void onSaveAsFile() {
//TODO save file
throw new UiException("save as file not implmented yet");
/*
		BufferedReader reader = null;
		BufferedWriter writer = null;
		try {
			// filename
			Textbox textbox = (Textbox) Path.getComponent("//p1/mainWin/fileSaveAsWin/fsaw_filename");
			String filename = textbox.getValue();

			// check duplicate filename
			getUniqueFilename(filename);

			ExcelExporter exporter = new ExcelExporter();
			long millis = System.currentTimeMillis();
			String hashFilename = millis + ".xls";
			
			exporter.exports(spreadsheet.getBook(), new File(xlsDir	+ hashFilename));
			
			Properties prop = new Properties();
			reader = new BufferedReader(new FileReader(xlsDir + "metaFile"));
			prop.load(reader);
			reader.close();
			reader = null;

			prop.put(filename,  millis + ", " + filename + ", "	+ hashFilename);
			writer = new BufferedWriter(new FileWriter(xlsDir + "metaFile", false));
			prop.store(writer, null);
			writer.close();
			writer = null;
			
			MainWindowCtrl w1 = MainWindowCtrl.getInstance();
			w1.hashFilename = hashFilename;

			Window fileOpenWin = (Window) Path.getComponent("//p1/mainWin/fileSaveAsWin");
			fileOpenWin.setVisible(false);

			Label label = (Label) Path.getComponent("//p1/mainWin/openingFilename");
			label.setValue(filename);
		} catch (Exception e) {
			e.printStackTrace();
		} finally {
			if (reader != null) {
				try {
					reader.close();
				} catch (IOException e) {
				}
			}
			if (writer != null) {
				try {
					writer.close();
				} catch (IOException e) {
				}
			}
		}
*/		
	}
	
	public String getUniqueFilename(String filename) {
		HashMap hm = readMetafile();
		try {
			while (hm.get(filename) != null) {
				filename = "duplicate_" + filename;
			}
		} catch (Exception e) {
			e.printStackTrace();
		}
		return filename;
	}

	public void saveFileInSS(String filename) {
//TODO remove me, save file		
throw new UiException("save file is not implemented yet");		
/*		ExcelExporter exporter = new ExcelExporter();
		String absPathFile = filename;//rootDir.concat(filename);
		File file = new File(absPathFile);
		exporter.exports(spreadsheet.getBook(), file);
*/	}

	public void onPopupMenu(String str){
		Window win = (Window) Path.getComponent(str);	
		win.setLeft("100px");
		win.setTop("100px");
		win.doPopup();
	}
	
	public void onOpenFileSelected() {
		reloadMenu("open");
		onPopupMenu("//p1/mainWin/fileOpenWin");
	}

	public void onSaveAsFileSelected() {
		onPopupMenu("//p1/mainWin/fileSaveAsWin");
	}

	public void onDeleteFileSelected() {
		reloadMenu("delete");
		onPopupMenu("//p1/mainWin/fileDeleteWin");
	}

	public void onExportFileSelected() {
		reloadMenu("export");
		onPopupMenu("//p1/mainWin/fileExportWin");
	}
	
	static public HashMap readMetafile(){
		BufferedReader bReader;
		HashMap hm = null;
		Properties prop = new Properties();
		try {
			bReader = new BufferedReader(new FileReader(xlsDir + "metaFile"));
			prop.load(bReader);
			
			prop.list(System.out);
			Iterator iter = prop.values().iterator();
			String timeStr, filename, hashFilename;
			long time;
			hm = new HashMap();
			Object objs[];
			while (iter.hasNext()) {
				String info = iter.next().toString();
				String[] setting = info.split(",");
				if (setting.length != 3)
					continue;
				
				timeStr = setting[0].trim();
				filename = setting[1].trim();
				hashFilename = setting[2].trim();
				
				time = Long.parseLong(timeStr);
				objs = (Object[]) hm.get(filename);
				if (objs == null || ((Long) objs[0]).longValue() < time) {
					hm.put(filename, new Object[] { new Long(time), hashFilename });
				}
			}
			/*
			String timeStr, filename, hashFilename;
			long time, timeTarget;
			hm = new HashMap();
			Object objs[];
			while (true) {
				timeStr = bReader.readLine();
				if (timeStr == null) {
					break;
				}
				filename = bReader.readLine();
				if (filename == null) {
					break;
				}
				hashFilename = bReader.readLine();
				if (hashFilename == null) {
					break;
				}
				time = Long.parseLong(timeStr);
				objs = (Object[]) hm.get(filename);
				if (objs == null || ((Long) objs[0]).longValue() < time) {
					hm.put(filename, new Object[] { new Long(time), hashFilename });
				}
			}
			*/
			bReader.close();
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		}
		return hm;
	}
	
	static public void deleteFile(String delFilename) {
		BufferedReader bReader = null;
		BufferedWriter bWriter = null;
		try {
			bReader = new BufferedReader(new FileReader(xlsDir + "metaFile"));
			Properties prop = new Properties();
			prop.load(bReader);
			bReader.close();
			bReader = null;
			
			if (prop.containsKey(delFilename)) {
				bWriter = new BufferedWriter(new FileWriter(xlsDir + "metaFile", false));
				prop.remove(delFilename);
				prop.store(bWriter, null);
				bWriter.close();
				bWriter = null;
			}
			
			/*
			String line = null;
			Queue aQueue = new LinkedList();

			String timeStr, filename, hashFilename;
			while (true) {
				timeStr = bReader.readLine();
				if (timeStr == null) {
					break;
				}
				filename = bReader.readLine();
				if (filename == null) {
					System.out.println("Warning: filename cannot read from metaFile");
					break;
				}
				hashFilename = bReader.readLine();
				if (hashFilename == null) {
					System.out.println("Warning: hashFilename cannot read from metaFile");
					break;
				}
				if (!delFilename.equals(filename)) {
					aQueue.add(timeStr);
					aQueue.add(filename);
					aQueue.add(hashFilename);
				} else {
					new File(xlsDir + hashFilename).delete();
				}
			}
			*/
			/*
			bWriter = new BufferedWriter(new FileWriter(xlsDir + "metaFile", false));
			while (!aQueue.isEmpty()) {
				bWriter.write(((String) aQueue.poll()) + "\n");
			}
			bWriter.close();
			*/
		} catch (FileNotFoundException e) {
			e.printStackTrace();
		} catch (IOException e) {
			e.printStackTrace();
		} finally {
			if (bReader != null) {
				try {
					bReader.close();
				} catch (IOException e) {
				}
			}
			if (bWriter != null) {
				try {
					bWriter.close();
				} catch (IOException e) {
				}
			}
		}

	}

	public void reloadMenu(String target) {
		HashMap hm = readMetafile();// the unique key(filename) value pair

		String filename, hashFilename;
		long time;
		Object objs[];
		Listbox listbox = null;
		if (target.equals("open"))
			listbox = (Listbox) Path.getComponent("//p1/mainWin/fileOpenWin/flo_files");
		if (target.equals("delete"))
			listbox = (Listbox) Path.getComponent("//p1/mainWin/fileDeleteWin/fld_files");
		if (target.equals("export"))
			listbox = (Listbox) Path.getComponent("//p1/mainWin/fileExportWin/fle_files");

		List childList = listbox.getChildren();
		// remain the list head component
		while (childList.size() > 1)
			listbox.removeChild((Component) childList.get(1));

		Iterator i = hm.entrySet().iterator();
		while (i.hasNext()) {
			Map.Entry me = (Entry) i.next();
			objs = (Object[]) me.getValue();
			time = ((Long) objs[0]).longValue();
			filename = (String) me.getKey();
			hashFilename = (String) objs[1];

			Listitem newListItem = new Listitem();

			newListItem.appendChild(new Listcell(filename));
			newListItem.appendChild(new Listcell("test user"));
			newListItem.appendChild(new Listcell(new Date(time).toString()));
			listbox.appendChild(newListItem);
		}
	}	
}
