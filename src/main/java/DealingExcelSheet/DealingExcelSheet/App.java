package DealingExcelSheet.DealingExcelSheet;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;

import org.apache.poi.*;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 * Hello world!
 *
 */
public class App 
{
    public static void main( String[] args )
    {
        System.out.println( "Hello World!" );
        App ap = new App();
        ap.readExcel();
        ap.writeExcel();
        
    }
    
    public void writeExcel(){

    	try {
    		
			FileInputStream fis = new FileInputStream("E:\\sample.xlsx");
			
			Workbook wb =  new XSSFWorkbook(fis);
			Sheet shit = wb.getSheet("some");
			
			
			int rowCount = shit.getLastRowNum();
			
			for(int i=0;i<rowCount;i++){
				Row rw = shit.getRow(i);
				for(int j=0;j<rw.getLastCellNum();j++){
					Cell ce = rw.getCell(j);
					ce.setCellValue("Inportia");
					
				}	
				
			}
			FileOutputStream fos = new FileOutputStream("E:\\sample.xlsx");
			wb.write(fos);
		wb.close();
		fis.close();	
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
		
 
    }
    
    public void readExcel(){
    	
    	try {
    		
			FileInputStream fis = new FileInputStream("E:\\sample.xlsx");
			
			Workbook wb =  new XSSFWorkbook(fis);
			Sheet shit = wb.getSheet("some");
			
			
			int rowCount = shit.getLastRowNum();
			
			for(int i=0;i<rowCount;i++){
				Row rw = shit.getRow(i);
				for(int j=0;j<rw.getLastCellNum();j++){
					Cell ce = rw.getCell(j);
					
					System.out.print(ce);
					System.out.print( " ");
				}	
				System.out.println();
			}
			wb.close();
			fis.close();
		} catch (Exception e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
    }
}
