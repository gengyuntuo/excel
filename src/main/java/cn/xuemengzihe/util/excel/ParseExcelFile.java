package cn.xuemengzihe.util.excel;

import java.io.Closeable;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.io.InputStream;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.apache.poi.ss.usermodel.Workbook;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

/**
 * <h1>解析Excel文件</h1>
 * <p>
 * 解析Excel文件
 * </p>
 * 
 * @author 李春
 * @time 2017年3月13日 下午3:07:46
 */
public class ParseExcelFile implements Closeable {
	private final Logger logger = LoggerFactory.getLogger(ParseExcelFile.class);
	/**
	 * 初始化状态
	 */
	public static final String STATUS_INIT = "init";
	/**
	 * 解析中状态
	 */
	public static final String STATUS_PARSE = "parse";
	/**
	 * 异常状态，该状态下无法正常解析
	 */
	public static final String STATUS_ERROR = "error";
	/**
	 * 完成解析，该状态下无法正常解析
	 */
	public static final String STATUS_FINISH = "finish";
	private Workbook workBook;
	private String parseStatus; // 初始化状态：init，解析状态：parse，异常状态：error
	private InputStream inputStream;

	{
		this.parseStatus = ParseExcelFile.STATUS_INIT; // 初始化
	}

	/**
	 * 获取Cell（单元格）中的内容，并返回其值（String类型）
	 * 
	 * @param cell
	 * @return
	 */
	@SuppressWarnings("deprecation")
	private String getCellValue(Cell cell) {
		String value = null;
		switch (cell.getCellType()) {
		case Cell.CELL_TYPE_BLANK:
			value = "";
			break;
		case Cell.CELL_TYPE_BOOLEAN:
			value = cell.getBooleanCellValue() ? "Y" : "N";
			break;
		case Cell.CELL_TYPE_ERROR:
			value = "";
			break;
		case Cell.CELL_TYPE_FORMULA:
			value = cell.getCellFormula();
			break;
		case Cell.CELL_TYPE_NUMERIC:
			value = Double.valueOf(cell.getNumericCellValue()).longValue() + "";
			break;
		case Cell.CELL_TYPE_STRING:
			value = cell.getStringCellValue();
			break;
		}
		return value;
	}

	/**
	 * 使用Excel的InputStream输入流构造
	 * 
	 * @param inputStream
	 */
	public ParseExcelFile(InputStream inputStream) {
		try {
			this.inputStream = inputStream;
			workBook = new HSSFWorkbook(inputStream);
			this.parseStatus = ParseExcelFile.STATUS_PARSE;
			logger.debug("Excel file init finished![ use InputStream]");
		} catch (IOException e) {
			logger.error("Excel file init failed![ use InputStream]");
			this.parseStatus = ParseExcelFile.STATUS_ERROR;
			try {
				this.close();
			} catch (IOException ee) {
			}
			e.printStackTrace();
		}
	}

	/**
	 * 使用Excel文件的File对象构造
	 * 
	 * @param file
	 * @throws IOException
	 */
	public ParseExcelFile(File file) throws IOException {
		try {
			inputStream = new FileInputStream(file);
			workBook = new HSSFWorkbook(inputStream);
			this.parseStatus = ParseExcelFile.STATUS_PARSE;
			logger.debug("Excel file init finish![ use File " + file.getName()
					+ "]");
		} catch (IOException e) {
			logger.error("Excel file init failed![ use File " + file.getName()
					+ "]");
			this.parseStatus = ParseExcelFile.STATUS_ERROR;
			try {
				this.close();
			} catch (IOException e1) {
			}
			e.printStackTrace();
			throw e;
		}
	}

	/**
	 * 获得当前文件的解析状态
	 * 
	 * @return
	 */
	public String getParseStatus() {
		return this.parseStatus;
	}

	/**
	 * 获得当前文件的中工作表的数量
	 * 
	 * @return
	 */
	public int getSheetNumber() {
		if (workBook != null) {
			return workBook.getNumberOfSheets();
		} else
			return 0;
	}

	/**
	 * 获取工作表标题
	 * 
	 * @param sheetNum
	 *            工作表序号
	 * @return 标题
	 */
	public String getSheetTitle(int sheetNum) {
		Sheet sheet = workBook.getSheetAt(sheetNum);
		Row row = sheet.getRow(0);
		if (row == null) {
			return "";
		}
		Cell cell = row.getCell(0);
		String value = getCellValue(cell);
		logger.debug("Excel: get sheet title ->" + value);
		return value;
	}

	/**
	 * 获取工作表的列标题名称
	 * 
	 * @param sheetNum
	 *            工作表序号
	 * @return 工作表列标题集合{有序}
	 */
	public List<String> getColumnNames(int sheetNum) {
		List<String> columnNames = new ArrayList<>();
		Sheet sheet = workBook.getSheetAt(sheetNum);
		Row row = sheet.getRow(1); // 默认第二行为列标题
		if (row == null)
			return columnNames;
		Cell cell = null;
		for (int i = 0; i < row.getLastCellNum(); i++) {
			cell = row.getCell(i);
			if (cell == null) // 遇到空的单元格就停止解析
				break;
			columnNames.add(getCellValue(cell));
		}
		logger.debug("Excel: get sheet column names ->" + columnNames);
		return columnNames;
	}

	/**
	 * 获取工作表内容（默认从第3行开始为内容数据）
	 * 
	 * @param sheetNum
	 *            工作表的序号， 从0开始
	 * @return 工作表内容
	 */
	@SuppressWarnings("deprecation")
	public List<Map<String, String>> getSheetContent(int sheetNum) {
		List<Map<String, String>> content = new ArrayList<>();
		Map<String, String> rowValue = null;
		List<String> colNames = this.getColumnNames(sheetNum);
		Sheet sheet = workBook.getSheetAt(sheetNum);
		Row row = null;
		Cell cell = null;
		// 遍历文档中的所有行
		logger.debug("工作表的行数：" + sheet.getLastRowNum());
		for (int i = 2; i < sheet.getLastRowNum(); i++) {
			row = sheet.getRow(i);
			if (row == null) // 如果改行为空，则跳过
				continue;
			rowValue = new HashMap<>();
			// 遍历每一列的值
			for (int j = 0; j < colNames.size(); j++) {
				cell = row.getCell(j);
				if (cell == null || cell.getCellType() == Cell.CELL_TYPE_BLANK) // 如果Cell为空，则舍弃改行
					continue;
				rowValue.put(colNames.get(j), getCellValue(cell));
			}
			content.add(rowValue);
		}
		logger.debug("Excel: get sheet content \n" + content);
		return content;
	}

	@Override
	public void close() throws IOException {
		this.parseStatus = ParseExcelFile.STATUS_FINISH;
		if (workBook != null)
			workBook.close();
		if (inputStream != null)
			inputStream.close();
	}

}
