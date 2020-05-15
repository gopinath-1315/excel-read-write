package read.com;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.text.SimpleDateFormat;
import java.util.Date;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.DateUtil;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class Prac1 {

	public static String main(int a,int b) throws IOException {
		String data="";
		File f = new File("C:\\Users\\Gopi\\Desktop\\Readwrite practice1.Xlsx");
		FileInputStream fi = new FileInputStream(f);
		Workbook w = new XSSFWorkbook(fi);
		Sheet sheet = w.getSheet("sheet1");
		int physicalNumberOfRows = sheet.getPhysicalNumberOfRows();
		System.out.println(physicalNumberOfRows);
		Row row = sheet.getRow(a);
		Cell cell = row.getCell(b);
		int cellType = cell.getCellType();
		if (cellType == 1) {
			data = cell.getStringCellValue();
			System.out.println(data);

		} else if (cellType == 0) {
			if (DateUtil.isCellDateFormatted(cell)) {

				Date dt = cell.getDateCellValue();
				SimpleDateFormat sf = new SimpleDateFormat("dd-mm-yyyy");
				data = sf.format(dt);
				System.out.println(data);

			} else {
				double nc = cell.getNumericCellValue();
				long l = (long) nc;
				data = String.valueOf(l);
				System.out.println(data);
			}

		}
		return data;

	}
}
