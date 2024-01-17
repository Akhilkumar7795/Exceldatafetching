import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.Iterator;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class DataDriven {

	public static void main(String[] args) throws IOException {

        FileInputStream fis = new FileInputStream("C://Users//AS086960//OneDrive - Cerner Corporation//testdata.xlsx");
       // C://Users//rahul//Documents//demodata.xlsx
		XSSFWorkbook  workbook = new XSSFWorkbook(fis);
		int sheets=workbook.getNumberOfSheets();
		for(int i=0;i<sheets;i++)
		{
			if(workbook.getSheetName(i).equalsIgnoreCase("testdata")) {
				XSSFSheet sheet=workbook.getSheetAt(i);
				//identify testcases column by scanning the entire 1st row
				Iterator<Row> rows=sheet.iterator();;//sheet is collection of rows
				Row firstrow=rows.next();
				Iterator<Cell> ce=firstrow.cellIterator();
				int k=0;
				int column=0;
				while(ce.hasNext())
				{
					Cell value=ce.next();
					if(value.getStringCellValue().equalsIgnoreCase("Data2")) {
						
						column=k;
					}
					k++;
				}
				System.out.println(column);
				//once column is identified then scan the entire column to identify the purchase test
				while(rows.hasNext())
				{
					Row r=rows.next();
					if(r.getCell(column).getStringCellValue().equalsIgnoreCase("purchase"))
					{
						//after you grab the purchase testcase row pull all the data of the row and feed in to test
						Iterator<Cell> cv=r.cellIterator();
						while(cv.hasNext()) {
							System.out.println(cv.next().getStringCellValue());
						}
					}
				}
			}
		}
		

	}

}
