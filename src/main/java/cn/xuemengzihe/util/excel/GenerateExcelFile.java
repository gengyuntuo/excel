package cn.xuemengzihe.util.excel;

import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.usermodel.Workbook;
import org.apache.poi.ss.util.CellRangeAddress;
import org.junit.Test;

/**
 * <h1>生成Excel文件</h1>
 * <p>
 * 生成Excel文件
 * </p>
 * 
 * @author 李春
 * @time 2017年2月23日 下午9:49:49
 */
public class GenerateExcelFile {
	private Workbook workbook;
	private String[] colsName;

	public GenerateExcelFile() {
		workbook = new HSSFWorkbook();
	}

	@Test
	public void generateExcelToLocalDir() throws IOException {
		FileOutputStream file1 = new FileOutputStream("workbook1.xls");
		FileOutputStream file2 = new FileOutputStream("workbook2.xls");
		Sheet sheet = workbook.createSheet();
		createTitleCell(sheet, "Title", 6);
		createColumnNames(sheet, new String[] { "列名1", "列名2", "列名3", "列名4",
				"列名5", "列名6" });
		String[] colsName = new String[] { "列名1", "列名2", "列名3", "列名4", "列名5",
				"列名6" };
		List<Map<String, String>> list = new ArrayList<>();
		for (int i = 0; i < 30; i++) {
			Map<String, String> map = new HashMap<String, String>();
			for (int j = 0; j < colsName.length; j++) {
				map.put(colsName[j], colsName[j] + "的值");
			}
			list.add(map);
		}
		createSheetContent(sheet, list);
		workbook.close();
		workbook.write(file1);
		workbook.write(file2);
		file1.close();
		file2.close();
	}

	/**
	 * 创建单元格的标题，这里使用的是默认位置（即单元格的第一行），样式：楷体 22号 垂直水平居中
	 * 
	 * @param sheet
	 *            表单
	 * @param title
	 *            标题
	 * @param colspan
	 *            占用列的宽度（单位：/个单元格）
	 * @return 被设置的表单
	 */
	public Sheet createTitleCell(Sheet sheet, String title, int colspan) {
		Row row = sheet.createRow(0);
		Cell cell = row.createCell(0);
		// 单元格字体美化
		Font font = workbook.createFont();
		font.setBold(true);
		font.setFontName("楷体");
		font.setFontHeightInPoints((short) 22);
		// 单元格单元格样式美化
		CellStyle style = workbook.createCellStyle();
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		cell.setCellValue(title);
		cell.setCellStyle(style);
		sheet.addMergedRegion(new CellRangeAddress(0, 0, 0, colspan - 1)); // 合并单元格
		return sheet;
	}

	/**
	 * 创建龚总表的列名
	 * 
	 * @param sheet
	 *            工作表
	 * @param colsName
	 *            列名数组
	 * @return 工作表
	 */
	public Sheet createColumnNames(Sheet sheet, String... colsName) {
		Row row = sheet.createRow(1);
		// 单元格字体美化
		Font font = workbook.createFont();
		font.setBold(true);
		font.setFontName("楷体");
		font.setFontHeightInPoints((short) 12);
		// 单元格样式美化
		CellStyle style = workbook.createCellStyle();
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		row.setRowStyle(style);
		// 创建列名
		Cell cell = null;
		for (int i = 0; i < colsName.length; i++) {
			cell = row.createCell(i);
			cell.setCellStyle(style);
			cell.setCellValue(colsName[i]);
		}
		// 保存列名，以便插入内容时使用
		this.colsName = colsName;
		return sheet;
	}

	/**
	 * 写入内容到工作表中,默认从Sheet的第三行开始写入数据（第一行为标题，第二行为列名称）
	 * 
	 * @param sheet
	 *            工作表
	 * @param content
	 *            内容
	 */
	public void createSheetContent(Sheet sheet,
			List<Map<String, String>> content) {
		// 1. 列名称判断
		if (this.colsName == null || this.colsName.length == 0) {
			throw new RuntimeException("列名不能为空！[can't find the column's name!]"); // 列名不能为空
		}
		// 2. 单元格样式
		Font font = workbook.createFont();
		font.setFontName("宋体");
		font.setFontHeightInPoints((short) 11);
		CellStyle style = workbook.createCellStyle();
		style.setFont(font);
		style.setAlignment(HorizontalAlignment.CENTER);
		style.setVerticalAlignment(VerticalAlignment.CENTER);
		style.setWrapText(true); // 适配单元格输出
		// 写入数据
		int i = 2; // 行标迭代变量，由于从第三行开始，所以初始值为2
		Row row = null;
		Cell cell = null;
		for (Map<String, String> rowData : content) {
			row = sheet.createRow(i++);
			for (int j = 0; j < this.colsName.length; j++) {
				cell = row.createCell(j);
				cell.setCellStyle(style);
				cell.setCellValue(rowData.get(this.colsName[j]));
			}
		}
	}
}
