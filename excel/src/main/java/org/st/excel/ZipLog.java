package org.st.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.time.LocalDateTime;
import java.util.Arrays;
import java.util.Iterator;
import java.util.List;
import java.util.zip.ZipFile;

import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import net.lingala.zip4j.io.inputstream.ZipInputStream;
import net.lingala.zip4j.io.outputstream.ZipOutputStream;
import net.lingala.zip4j.model.LocalFileHeader;
import net.lingala.zip4j.model.ZipParameters;
import net.lingala.zip4j.model.enums.AesKeyStrength;
import net.lingala.zip4j.model.enums.CompressionMethod;
import net.lingala.zip4j.model.enums.EncryptionMethod;

import java.util.List;

public class ZipLog {
	
	String printstation_id;
	
	ZipLog (String id) throws IOException {
		String excel = (printstation_id = id) + "-log.xlsx";
				
		if ( (new File (excel)).exists () )
			return;
		
		XSSFWorkbook myExcelBook = new XSSFWorkbook (); //OPCPackage.open (new File (file)));
		myExcelBook.createSheet ("Events Log");
    	
     	// return first sheet from the XLSX workbook
    	// XSSFSheet mySheet = myExcelBook.getSheetAt (0);
    	FileOutputStream out = new FileOutputStream (new File (excel));
		
    	myExcelBook.write (out);
		out.close();
		
		zipOutputStream (new File (printstation_id + ".zlog"), Arrays.asList (new File (excel)), 
						 "lesing".toCharArray (),
						 CompressionMethod.STORE,
						 true,
						 EncryptionMethod.AES, AesKeyStrength.KEY_STRENGTH_256); 
		
		System.out.println ("Done");
		}

	public ZipLogEvent last () throws IOException {
		zipInputStream (new File (printstation_id + ".zlog"), "lesing".toCharArray ());
		
		ZipLogEvent result = null;
		
		FileInputStream input = new FileInputStream (new File (printstation_id + "-log.xlsx")); 
			
		XSSFWorkbook myExcelBook = new XSSFWorkbook (input);
		XSSFSheet sheet = myExcelBook.getSheetAt (0);
			
		if ( sheet.getLastRowNum () == -1 )
			return result;
			
		Row r = sheet.getRow (sheet.getLastRowNum ());
	    result = new ZipLogEvent (r.getCell(0).getLocalDateTimeCellValue(), r.getCell (1).getStringCellValue (), 
	    						  											r.getCell (2).getStringCellValue (),
	    						  											r.getCell (3).getStringCellValue (),
	    						  											r.getCell (4).getStringCellValue ()
	    						  											);
			
	    myExcelBook.close ();
		 
		return result;
		}
			
	public void add (String _place, String _room, String _event, String _description) throws IOException {
		zipInputStream (new File (printstation_id + ".zlog"), "lesing".toCharArray ());
		
		FileInputStream input = new FileInputStream (new File (printstation_id + "-log.xlsx")); 
		
		XSSFWorkbook myExcelBook = new XSSFWorkbook (input);
		XSSFSheet sheet = myExcelBook.getSheetAt (0);
			
		Row r = sheet.createRow (sheet.getLastRowNum () + 1);
			
		r.createCell (0).setCellValue (LocalDateTime.now ());
		r.createCell (1).setCellValue (_place);
		r.createCell (2).setCellValue (_room);
		r.createCell (3).setCellValue (_event);
		r.createCell (4).setCellValue (_description);
			
		FileOutputStream outputStream = new FileOutputStream (printstation_id + "-log.xlsx");
		
		myExcelBook.write (outputStream);
		myExcelBook.close ();
        
		outputStream.close ();
		
		zipOutputStream (new File (printstation_id + ".zlog"), Arrays.asList (new File (printstation_id + "-log.xlsx")), 
				 "lesing".toCharArray (),
				 CompressionMethod.STORE,
				 true,
				 EncryptionMethod.AES, AesKeyStrength.KEY_STRENGTH_256); 
		}

	private void zipOutputStream (File outputZipFile, List <File> filesToAdd, 
								 char [] password,  
             					 CompressionMethod compressionMethod, boolean encrypt,
             					 EncryptionMethod encryptionMethod, AesKeyStrength aesKeyStrength) throws IOException {

		ZipParameters zipParameters = buildZipParameters (compressionMethod, encrypt, encryptionMethod, aesKeyStrength);
		 
		byte [] buff = new byte [4096];
		int readLen;

		try ( ZipOutputStream zos = initializeZipOutputStream (outputZipFile, encrypt, password) ) {
			for ( File fileToAdd : filesToAdd ) {
				// Entry size has to be set if you want to add entries of STORE compression method (no compression)
				// This is not required for deflate compression
				if ( zipParameters.getCompressionMethod () == CompressionMethod.STORE ) {
					zipParameters.setEntrySize (fileToAdd.length ());
				 	}

				zipParameters.setFileNameInZip (fileToAdd.getName ());
				zos.putNextEntry (zipParameters);

				try ( InputStream inputStream = new FileInputStream (fileToAdd) ) {
					while ( (readLen = inputStream.read (buff)) != -1 ) {
						zos.write(buff, 0, readLen);
					 	}
				 	}
				zos.closeEntry ();
			 	}
			}
		}

	private ZipOutputStream initializeZipOutputStream (File outputZipFile, boolean encrypt, char [] password) throws IOException {
		FileOutputStream fos = new FileOutputStream (outputZipFile);

		if ( encrypt ) {
			return new ZipOutputStream (fos, password);
		 	}

		return new ZipOutputStream(fos);
	 	}

	private ZipParameters buildZipParameters(CompressionMethod compressionMethod, boolean encrypt,
		 								     EncryptionMethod encryptionMethod, AesKeyStrength aesKeyStrength) {
		ZipParameters zipParameters = new ZipParameters ();
		 
		zipParameters.setCompressionMethod(compressionMethod);
		zipParameters.setEncryptionMethod(encryptionMethod);
		zipParameters.setAesKeyStrength(aesKeyStrength);
		zipParameters.setEncryptFiles(encrypt);
    
		return zipParameters;
	 	}
	
	private void zipInputStream (File inputZipFile, char [] password) throws IOException {
		LocalFileHeader localFileHeader;
		    
		int readLen;
		byte [] readBuffer = new byte [4096];

		InputStream inputStream = new FileInputStream (inputZipFile);
		
		try ( ZipInputStream zipInputStream = new ZipInputStream (inputStream, password) ) {
			while ( (localFileHeader = zipInputStream.getNextEntry ()) != null ) {
		        File extractedFile = new File (localFileHeader.getFileName ());
		        
		        try ( OutputStream outputStream = new FileOutputStream (extractedFile) ) {
		        	while ( (readLen = zipInputStream.read (readBuffer) ) != -1 )
		        		outputStream.write (readBuffer, 0, readLen);
		        	}
				}
			}
		}
	}
	