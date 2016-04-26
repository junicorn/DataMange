package com.file.util;

import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.zip.ZipEntry;
import java.util.zip.ZipOutputStream;

public class FileZip {
	public static void ZipFiles(java.io.File[] srcfile, java.io.File zipfile) {  
		byte[] buf = new byte[1024];  
		try {  
			ZipOutputStream out = new ZipOutputStream(new FileOutputStream(zipfile));  
			for (int i = 0; i < srcfile.length; i++) {  
				FileInputStream in = new FileInputStream(srcfile[i]);  
				out.putNextEntry(new ZipEntry(srcfile[i].getName()));  
				int len;  
				while ((len = in.read(buf)) > 0) {  
					out.write(buf, 0, len);  
				}  
				out.closeEntry();  
				in.close();  
			}  
			out.close();  
		} catch (IOException e) {  
			e.printStackTrace();  
		}  
	}  
}
