package ActionClass;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class example4fortable {


		public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
			
			File myfile=new File("C:\\Users\\ASUS\\Documents\\Book1.xlsx");
			Workbook mybook = WorkbookFactory.create(myfile);
			Sheet type = mybook.getSheet("Sheet1");
			//for reading multiple data from single row
			for(int i =0;i<=2;i++) {
				String val = type.getRow(0).getCell(i).getStringCellValue();
				System.out.println(val);
				
			}
			System.out.println("**************************************");
			//reading multiple data from single column
			for(int i = 0;i<=2;i++) {
				String val = type.getRow(i).getCell(0).getStringCellValue();
				System.out.println(val);
				
			}

	}

}
