package cn.xuemengzihe.util.excel;

import java.io.File;
import java.io.IOException;

import org.junit.Test;

public class RunTest {

	@Test
	public void test() {
		File file = new File("workbook3.xls");
		ParseExcelFile excel = new ParseExcelFile(file);
		System.out.println(excel.getSheetTitle(0));
		System.out.println(excel.getColumnNames(0));
		System.out.println(excel.getSheetContent(0));
		try {
			excel.close();
		} catch (IOException e) {
		}
	}

}
