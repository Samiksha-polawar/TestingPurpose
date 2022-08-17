package ActionClass;

import java.io.File;
import java.io.IOException;


import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class example1 {

	
		

			
			public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
						
						File myfile=new File("C:\\Users\\ASUS\\Documents\\Book1.xlsx");
					
						String l = WorkbookFactory.create(myfile).getSheet("Sheet1").getRow(0).getCell(0).getStringCellValue();
System.out.println(l);
String l1 = WorkbookFactory.create(myfile).getSheet("Sheet1").getRow(0).getCell(1).getStringCellValue();
System.out.println(l1);
String l2 = WorkbookFactory.create(myfile).getSheet("Sheet1").getRow(0).getCell(2).getStringCellValue();
System.out.println(l2);
			}

}
