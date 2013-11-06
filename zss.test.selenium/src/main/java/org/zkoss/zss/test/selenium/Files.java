package org.zkoss.zss.test.selenium;

import java.awt.image.BufferedImage;
import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.io.Reader;
import java.io.StringWriter;
import java.io.Writer;

import javax.imageio.ImageIO;

public class Files {

	public static void copy(InputStream from, OutputStream to) throws IOException {
		byte[] buff = new byte[1024];
		int r;
		while ((r = from.read(buff)) != -1) {
			to.write(buff, 0, r);
		}
	}

	public static void copy(File from, File to) throws IOException {
		FileInputStream fis = new FileInputStream(from);
		if(!to.exists()){
			to.createNewFile();
		}
		FileOutputStream fos = new FileOutputStream(to);

		try {
			copy(fis, fos);
		} finally {
			if (fis != null) {
				try {
					fis.close();
				} catch (Exception x) {
				}
			}
			if (fos != null) {
				try {
					fos.close();
				} catch (Exception x) {
				}
			}
		}
	}
	
	public static BufferedImage readImage(File file) throws IOException{
		BufferedImage img = ImageIO.read(file);
		return img;
	}
	
	public static void writeImage(File file,BufferedImage img) throws IOException{
		ImageIO.write(img,"png",file);
	}
	
	public static StringBuffer readAll(Reader reader) throws IOException{
		final StringWriter writer = new StringWriter(1024*16);
		copy(writer, reader);
		return writer.getBuffer();
	}

	public static byte[] readAll(InputStream in) throws IOException{
		final ByteArrayOutputStream out = new ByteArrayOutputStream();
		final byte[] buf = new byte[1024*16];
		for (int v; (v = in.read(buf)) >= 0;) { //including 0
			out.write(buf, 0, v);
		}
		return out.toByteArray();
	}
	
	public static final void copy(Writer writer, Reader reader)
	throws IOException {
		final char[] buf = new char[1024*4];
		for (int v; (v = reader.read(buf)) >= 0;) {
			if (v > 0)
				writer.write(buf, 0, v);
		}
	}
}
