import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;

import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.IndexedColors;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POI_Function {
	public static String filePath = "D:\\D_other";
	public static String fileNm = "char.xlsx";

	public static void main(String[] args) {
		try {
			XSSFWorkbook workbook = new XSSFWorkbook();
			XSSFSheet sheet = workbook.createSheet("Sheet1");
			
			

			CellStyle style = workbook.createCellStyle();
			style.setFillForegroundColor(IndexedColors.AQUA.getIndex());
			style.setFillPattern(FillPatternType.FINE_DOTS);
			style.setBorderTop(BorderStyle.DOTTED);
			style.setBorderBottom(BorderStyle.HAIR);
			style.setBorderLeft(BorderStyle.THICK);
			style.setBorderRight(BorderStyle.MEDIUM);
			
			Row row1 = sheet.createRow(0);
			Cell cellA1 = row1.createCell(0);
			Cell cellB1 = row1.createCell(1);

			// 셀에 값 설정
			cellA1.setCellValue(10);
			cellB1.setCellValue(20);

			// SUM 함수 사용
			Row row2 = sheet.createRow(1);
			Cell cellA2 = row2.createCell(0);
			cellA2.setCellValue("SUM(A1:B1)");
			Cell cellA3 = row2.createCell(1);
			cellA3.setCellFormula("SUM(A1:B1)");
			
			
			// AVERAGE 함수 사용
			Row row3 = sheet.createRow(2);
			Cell cellA4 = row3.createCell(0);
			cellA4.setCellValue("AVERAGE(A1:B1)");
			Cell cellA5 = row3.createCell(1);
			cellA5.setCellFormula("AVERAGE(A1:B1)");

			// FormulaEvaluator를 사용하여 수식 계산
			FormulaEvaluator evaluator = workbook.getCreationHelper().createFormulaEvaluator();
			evaluator.evaluateFormulaCell(cellA3);
			evaluator.evaluateFormulaCell(cellA5);

			
			Cell cellColor = sheet.createRow(5).createCell(0); 
			CellRangeAddress region = CellRangeAddress.valueOf("A6:C7");
			sheet.addMergedRegion(region);
			cellColor.setCellStyle(style);

			// 파일 저장
			FileOutputStream outputStream = new FileOutputStream(new File(filePath,fileNm));
			workbook.write(outputStream);
			workbook.close();
			outputStream.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
