import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class poi_basic {
	public static String filePath = "D:\\D_other";
	public static String fileNm = "WriteExample.xlsx";

	public static void main(String[] args) {

		XSSFWorkbook workbook = new XSSFWorkbook();
		XSSFSheet sheet = workbook.createSheet("대덕인재개발원");
		XSSFSheet sheet2 = workbook.createSheet("대덕인재개발원2");
		XSSFRow xrow = null;
		Cell cell = null;

		CellStyle style = workbook.createCellStyle();
		style.setFillForegroundColor(IndexedColors.LIGHT_GREEN.getIndex());
		style.setFillPattern(FillPatternType.FINE_DOTS);
		style.setBorderTop(BorderStyle.DOTTED);
		style.setBorderBottom(BorderStyle.HAIR);
		style.setBorderLeft(BorderStyle.THICK);
		style.setBorderRight(BorderStyle.MEDIUM);
		
		xrow = sheet.createRow(0);
		cell = xrow.createCell(0);
		cell.setCellValue("1번!");

		xrow = sheet.createRow(1);
		cell = xrow.createCell(2);
		cell.setCellValue("2번!");



		for (int i = 3; i < 10; i++) {
			xrow = sheet.createRow(i);
			for (int j = 3; j < 10; j++) {
				cell = xrow.createCell(j);
				cell.setCellValue(i);
				cell.setCellStyle(style);
			}
			
		}

		try {
			FileOutputStream out = new FileOutputStream(new File(filePath, fileNm));
			workbook.write(out);
			out.close();
		} catch (IOException e) {
			// TODO: handle exception
		}

	}

}
