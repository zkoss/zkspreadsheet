package org.zkoss.zss.ngapi.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.io.Serializable;

import org.zkoss.zss.ngapi.NExporter;
import org.zkoss.zss.ngapi.NImporter;
import org.zkoss.zss.ngmodel.NBook;

public abstract class AbstractExporter implements NExporter,Serializable{
	@Override
	public void export(NBook book, File file) throws IOException {
		OutputStream os = null;
		try{
			os = new FileOutputStream(file);
			export(book,os);
		}finally{
			if(os!=null){
				try{
					os.close();
				}catch(Exception x){};
			}
		}
	}
}
