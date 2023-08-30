package ExcelRead;
import java.io.FileInputStream;
import java.io.IOException;

import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Read1 {
	public static void main(String[] args) throws IOException  {
		String excelfilepath="C:\\Users\\Admin\\eclipse-workspace\\WCSM11sel-master.zip_expanded\\WCSM11sel-master\\data.xlsx";
		FileInputStream fi = new FileInputStream(excelfilepath);
		XSSFWorkbook workbook = new XSSFWorkbook(fi);
		XSSFSheet sheet = workbook.getSheet("Sheet1");		
		int rows=sheet.getLastRowNum(); // no of rows  rows count start with 0
		System.out.println(rows);
		int col=sheet.getRow(1).getLastCellNum();  // no of columns   colums count strat with 1 
		System.out.println(col);
		for(int r=0; r<=rows; r++) // for rows
		{
			XSSFRow row=sheet.getRow(r);
			for(int c=0; c<col; c++)  // for columns
			{
				XSSFCell cell=row.getCell(c);
			    switch(cell.getCellType())
			    {
			    case STRING : System.out.println(cell.getStringCellValue());
			    break;
			    case NUMERIC: System.out.println(cell.getNumericCellValue());
			    break;
			    case BOOLEAN: System.out.println(cell.getBooleanCellValue());
			    break;
			    }
			}
			System.out.println();
	}
  }
}
