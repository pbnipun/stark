package excelhtml.com.excelhtml;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.util.List;

import org.apache.commons.io.FileUtils;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;

public class Excelhtml {
	static String title;
	static String description;
	static String remediation;
	static String reference;
	static String result;
	static StringBuffer doc = new StringBuffer();

	public static void main(String[] args) throws Exception {
		// TODO Auto-generated method stub

		File src = new File("input.xls");

		FileInputStream fis = new FileInputStream(src);

		HSSFWorkbook wb = new HSSFWorkbook(fis);

		HSSFSheet sheet1 = wb.getSheetAt(0);

		String[] header = new String[4];
		for (int k = 0; k < 4; k++) {
			header[k] = sheet1.getRow(0).getCell(k).getStringCellValue();
		}

		for (int j = 1; j < 5; j++) {
			StringBuffer rowBuffer = new StringBuffer();
			for (int i = 0; i < 4; i++) {

				result = sheet1.getRow(j).getCell(i).getStringCellValue();
				rowBuffer.append("<b>" + header[i] + "</b>" + " : " + result + "<br>");

				System.out.println(doc.toString());

			}
			
			doc.append("<p>" + rowBuffer.toString() + "</p>");

			File newHtmlFile = new File("output.html");
			FileUtils.writeStringToFile(newHtmlFile,doc.toString());
			System.out.println();

		}
		wb.close();
	}
}
