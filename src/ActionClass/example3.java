package ActionClass;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class example3 {


	public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
				
				File myfile=new File("C:\\Users\\ASUS\\Documents\\Book1.xlsx");
				Workbook mybook = WorkbookFactory.create(myfile);
				Sheet type = mybook.getSheet("Sheet1");
				System.out.println("********************this is for numeric*************");
			Cell cellno = type.getRow(1).getCell(0);
			System.out.println(cellno.getCellType());
			System.out.println(cellno.getNumericCellValue());
			System.out.println("**************this is for string **********");
			Cell cellno1=type.getRow(0).getCell(0);
			System.out.println(cellno1.getCellType());
			System.out.println(cellno1.getStringCellValue());
			System.out.println("******this is for boolean *********************");
			Cell cellno2=type.getRow(2).getCell(0);
			System.out.println(cellno2.getCellType());
		
	}

}
