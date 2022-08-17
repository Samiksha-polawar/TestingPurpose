package ActionClass;

import java.io.File;
import java.io.IOException;

import org.apache.poi.EncryptedDocumentException;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;

import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.WorkbookFactory;

public class ex7 {

public static void main(String[] args) throws InterruptedException, EncryptedDocumentException, IOException {
		
		File myfile=new File("C:\\Users\\ASUS\\Documents\\Book1.xlsx");
		Sheet mybook = WorkbookFactory.create(myfile).getSheet("Sheet1");
		int rows = mybook.getLastRowNum();
		int col = mybook.getRow(0).getLastCellNum()-1;
		for(int i=0;i<=rows;i++) {
			
			
			for(int j=0;j<=col;j++)
			{
				Cell mycell = mybook.getRow(i).getCell(j);
			CellType datatype = mycell.getCellType();
			System.out.print(datatype+" || ");
			}
			System.out.println();
			
	
		}
	}

}
