package org.company;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.math.BigDecimal;
import java.text.SimpleDateFormat;
import java.util.Date;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
public class Datasheet {
	public static void main(String[] args) throws IOException {
File file = new File("C:\\Users\\malat\\eclipse-workspace\\excel\\excelfold\\date.xlsx");
FileInputStream stream = new FileInputStream(file);
Workbook W = new XSSFWorkbook(stream);
Sheet sheet = W.getSheet("sheet1");
//Row row =sheet.getRow(2);
for(int i=0;i<sheet.getPhysicalNumberOfRows();i++) {
	Row row1 =sheet.getRow(i);
   for(int j=0;j<row1.getPhysicalNumberOfCells();j++)	{
	Cell cell1 = row1.getCell(j);
	//System.out.println(cell1);
CellType cellType = cell1.getCellType();
switch(cellType)
{
	case STRING:
		String s = cell1.getStringCellValue();
		System.out.print(s+"\t");
		break;
	case NUMERIC:
		if(DateUtil.isCellDateFormatted(cell1)){
			Date datecellvalue= cell1.getDateCellValue();
			SimpleDateFormat dateformat = new SimpleDateFormat("DD/MMM/YY");
			String format = dateformat.format(datecellvalue);
			System.out.print(format+"\t");
		}
     else 
		{
			double numericCellValue = cell1.getNumericCellValue();
			//System.out.println(numericCellValue);
			long l = (long)numericCellValue;
			System.out.print(l+"\t");
		}
		break;
		default :
			break;
		
}
   }
   System.out.println();
   System.out.println("Sam Project");
   System.out.println("Usha Project");
}
}
}