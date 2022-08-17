package ActionClass;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ex6 {

public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
		//dynamic reading for complete excel sheet
		File myfile=new File("C:\\Users\\ASUS\\Documents\\Book1.xlsx");
		Sheet mybook = WorkbookFactory.create(myfile).getSheet("Sheet1");
	int rows = mybook.getLastRowNum();
	System.out.println("total no of rows "+rows);
	int col = mybook.getRow(0).getLastCellNum()-1;
	System.out.println("total no of columns" +col);
	for(int i=0;i<=rows;i++) {
		for(int j=0;j<=col;j++)
		{
			String val =mybook.getRow(i).getCell(j).getStringCellValue();
			System.out.print(val+" ||");
			
		}
		System.out.println();
	}
	
		
	}

}
