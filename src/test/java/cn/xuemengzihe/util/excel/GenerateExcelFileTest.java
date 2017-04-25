package cn.xuemengzihe.util.excel;

import java.io.IOException;

public class GenerateExcelFileTest {

	// @Test
	public void testGenerateExcelFile() {
		GenerateExcelFile gen = new GenerateExcelFile();
		gen.writeAndMerge("沈阳理工大学13030504班级测评成绩", 10, 2);
		gen.writeValue("nihao");
		gen.switchToNextRow();
		gen.writeValue("123456");
		gen.generateExcelFile("C:\\Users\\Lenovo\\Desktop\\test.xls");
		try {
			gen.close();
		} catch (IOException e) {
		}
	}

}
