package ActionClass;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;

import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class Example2 {

	
		public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
			
			File myfile=new File("C:\\Users\\ASUS\\Documents\\Book1.xlsx");
		
			Workbook mybook = WorkbookFactory.create(myfile);
			
			Sheet sht = mybook.getSheet("Sheet1");
			
			 Row myrow = sht.getRow(0);
			 
			Cell mycell = myrow.getCell(2);
			
		String value = mycell.getStringCellValue();
			System.out.println(value);
			System.out.println(sht.getRow(0).getCell(0).getStringCellValue());
					
			

	}

}
