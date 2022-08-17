package ActionClass;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Sheet;

import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ex5 {

	public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
		
		File myfile=new File("C:\\Users\\ASUS\\Documents\\Book1.xlsx");
		Sheet mybook = WorkbookFactory.create(myfile).getSheet("Sheet1");
		//for reading complete sheet
		for(int i=0;i<=2;i++) {
			for(int j=0;j<=2;j++)
			{
				String val =mybook.getRow(i).getCell(j).getStringCellValue();
				System.out.print(val+" ||");
				
			}
			System.out.println();
		}
		

}
}
