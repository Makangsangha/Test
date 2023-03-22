import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.Map;
import java.util.Set;
import java.util.TreeMap;

import org.apache.commons.collections4.map.HashedMap;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class POI_PROJECT {
	public static String filePath = "D:\\D_other";
	public static String fileNm = "WriteExample2.xlsx";

	public static void main(String[] args) {

		// 빈 Workbook 생성
		XSSFWorkbook workbook = new XSSFWorkbook();

		// 빈 Sheet를 생성
		XSSFSheet sheet1 = workbook.createSheet("대덕인재개발원");
		XSSFSheet sheet2 = workbook.createSheet("대덕인재개발원2");

		// Sheet를 채우기 위한 데이터들을 Map에 저장
		Map<String, Object[]> data = new TreeMap<>();
		data.put("1", new Object[] { "ID", "NAME", "PHONE_NUMBER" });
		data.put("2", new Object[] { "1", "김민욱", "010-1111-1111" });
		data.put("3", new Object[] { "2", "진현성", "010-2222-2222" });
		data.put("4", new Object[] { "3", "신국현", "010-3333-3333" });
		data.put("5", new Object[] { "4", "오대환", "010-4444-4444" });
		
		Map<String, Object[]> data2 = new HashedMap<>();
		data2.put("1", new Object[] { "ID", "NAME", "PHONE_NUMBER" });
		data2.put("2", new Object[] { "1", "정재균", "010-1111-1111" });
		data2.put("3", new Object[] { "2", "김동혁", "010-2222-2222" });
		data2.put("4", new Object[] { "3", "진현순", "010-3333-3333" });
		data2.put("5", new Object[] { "4", "진현돌", "010-4444-4444" });

		// data에서 keySet를 가져온다. 이 Set 값들을 조회하면서 데이터들을 sheet에 입력한다.
		Set<String> keyset = data.keySet();
		int rownum = 0;
	
		// 알아야할 점, TreeMap을 통해 생성된 keySet는 for를 조회시, 키값이 오름차순으로 조회된다.
		for (String key : keyset) {
			Row row = sheet1.createRow(rownum++);
			Object[] objArr = data.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof Integer) {
					cell.setCellValue((Integer) obj);
				}
			}
		}
		
		Set<String> keyset2 = data2.keySet();
		int rownum2 = 0;
		for (String key : keyset2) {
			Row row = sheet2.createRow(rownum2++);
			Object[] objArr = data2.get(key);
			int cellnum = 0;
			for (Object obj : objArr) {
				Cell cell = row.createCell(cellnum++);
				if (obj instanceof String) {
					cell.setCellValue((String) obj);
				} else if (obj instanceof Integer) {
					cell.setCellValue((Integer) obj);
				}
			}
		}
		try {
			FileOutputStream out = new FileOutputStream(new File(filePath, fileNm));
			workbook.write(out);
			out.close();
		} catch (IOException e) {
			e.printStackTrace();
		}

	}

}
