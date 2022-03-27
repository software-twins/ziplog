package org.st.excel;

import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.Iterator;

import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.WorkbookFactory;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        
        try { 
        	//readFromExcel ("read-test.xlsx");
        	ZipLog ziplog = new ZipLog ("test-ststion");
        	ziplog.add ("sdgfhffhhfd", Integer.toString (1), "ggygyg,gyg", "nknljnljljjnljjl");
        	System.out.print (ziplog.last ().timestamp);
        	} 
        catch (Exception e) {
            System.out.println(e);
            }
    }
    
    public static void readFromExcel(String file) throws IOException, InvalidFormatException {
    	
    	XSSFWorkbook myExcelBook = new XSSFWorkbook (OPCPackage.open (new File (file)));
    	
     	// Return first sheet from the XLSX workbook
    	XSSFSheet mySheet = myExcelBook.getSheetAt (0);
    	
    	 // Get iterator to all the rows in current sheet
    	 Iterator < Row > rowIterator = mySheet.iterator ();
    	 
    	 // Traversing over each row of XLSX file
    	 while ( rowIterator.hasNext () ) {
    		Row row = rowIterator.next();
      		System.out.println ("-");
    	 	}
    	 
    	 Row r = mySheet.getRow (mySheet.getLastRowNum ());
    	 
    	 int firstCell = r.getFirstCellNum ();
    	 
    	 for ( int i = 0; i < 5; i ++)
    		 System.out.print(r.getCell (i + firstCell).getNumericCellValue());
    	 
    	 System.out.println ( "Hello World! 3" );
     
      //  myExcelBook.close ();
    	}
	}
