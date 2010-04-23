package org.zkoss.zss.app;

import org.zkoss.zul.Messagebox;

public class MenuFile {

	public static void onFileNew() {
		
	}

	public static void onFileOpen() {
	}

	public static void onFileSave(){
		try {
			Messagebox.show("dummy");
		} catch (Exception e) {
			System.out.println(e);
		}		
	}

	public static void onFileSaveAs() {
	}

	public static void onFilePrintSetup() {
	}

	public static void onFilePrintRegion() {
	}

	public static void onFilePrintPDF() {
	}	
}
