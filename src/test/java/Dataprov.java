
	import java.io.FileInputStream;
	import java.io.FileNotFoundException;
	import java.io.IOException;

import org.apache.poi.ss.usermodel.DataFormatter;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.testng.annotations.DataProvider;
	import org.testng.annotations.Test;

	public class Dataprov {

		DataFormatter formatter = new DataFormatter();
		@Test(dataProvider="driveTest")
		public void testcasedata(String greeting,String communication,String id) {
		
			System.out.println(greeting+communication+id);
		}
		@DataProvider(name="driveTest")
		public Object[][] getData() throws IOException {
			
			//Object[][] data= {{"hello","text","1"},{"bye","cu","2"},{"solo","call","3"}};
			FileInputStream fis = new FileInputStream("C:\\Users\\AS086960\\OneDrive - Cerner Corporation\\Documents\\excelDriven.xlsx");
			XSSFWorkbook wb = new XSSFWorkbook();
			XSSFSheet sheet=wb.getSheetAt(0);
			int rowCount=sheet.getPhysicalNumberOfRows();
			XSSFRow row= sheet.getRow(0);
			int colCount=row.getLastCellNum();
			Object data[][]=new Object[rowCount-1][colCount];
			for(int i=0;i<rowCount-1;i++) {
				row =sheet.getRow(i+1);
				for(int j=0;j<colCount;j++) {
					XSSFCell cell=row.getCell(j);
					data[i][j]=formatter.formatCellValue(cell);
				}
			}
					
			return data;
		}
	}


