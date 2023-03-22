import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class workBookWrite2 {

	public static void main(String[] args) {
		String filePath = "D:\\D_Other";
		String fileName = "workBookWrite.xlsx";

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("대덕인재개발원2");
		XSSFRow xrow = sheet.createRow(0);
		Cell cell = xrow.createCell(0);
		cell.setCellValue("완성");

		sheet.createRow(3).createCell(5).setCellValue("완성2");

		try {
			FileOutputStream out = new FileOutputStream(new File(filePath, fileName));
			workbook.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
