package abcd;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class C {

	public static void main(String[] args) throws Exception {
		
  
	    FileInputStream file = new FileInputStream(new File("E:\\Book13.xlsx"));
        XSSFWorkbook yourworkbook = new XSSFWorkbook(file);
        XSSFSheet sheet1 = yourworkbook.getSheetAt(0);
        Row row = sheet1.getRow(1);
        Cell column = row.getCell(0);
        String updatename = column.getStringCellValue();
        updatename="Lala";
        column.setCellValue(updatename);
        System.out.println(updatename);
        file.close();
        FileOutputStream out = 
            new FileOutputStream(new File("E:\\Book13.xlsx"));
        yourworkbook.write(out);
        out.close();
	
	}
	
}
