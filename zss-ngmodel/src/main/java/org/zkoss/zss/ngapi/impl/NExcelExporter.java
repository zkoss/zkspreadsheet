package org.zkoss.zss.ngapi.impl;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;

import org.zkoss.zss.ngapi.NExporter;
import org.zkoss.zss.ngapi.impl.ExcelExportFactory.Type;
import org.zkoss.zss.ngmodel.NBook;

public class NExcelExporter implements NExporter{
	
	private final ExcelExportFactory.Type type;
	
	public NExcelExporter(Type type){
		this.type = type;
	}
	
	@Override
	public void export(NBook book, File file) throws IOException {
		FileOutputStream fos = null;
		try{
			export(book,fos);
		}finally{
			if(fos!=null){
				try{
					fos.close();
				}catch(Exception x){};
			}
		}
	}
	
	@Override
	public void export(NBook book, OutputStream fos) throws IOException {
		// TODO Auto-generated method stub
		
	}

}
