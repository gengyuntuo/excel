package cn.xuemengzihe.util.excel;

import java.io.Closeable;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.ss.usermodel.VerticalAlignment;
import org.apache.poi.ss.util.CellRangeAddress;

/**
 * 
 * <h1>生成Excel文档</h1>
 * <p>
 * <b>注意：</b>线程不安全
 * </p>
 * 
 * @author 李春
 * @time 2017年4月24日 上午8:52:05
 */
public class GenerateExcelFile implements Closeable {
	private HSSFWorkbook workbook;
	private String[] colsName;
	private HSSFSheet curWorkSheet; // 当前的工作表
	private HSSFRow row; // 行
	private HSSFCell cell; // 单元格
	private HSSFFont font; // 当前的字体
	private HSSFCellStyle style; // 当前的样式
	private int curRow; // 当前行
	private int curColumn; // 当前列
	private boolean isClosed = false;

	public GenerateExcelFile() {
		this.workbook = new HSSFWorkbook();
		this.curWorkSheet = workbook.createSheet();
		this.curRow = 0;
		this.curColumn = 0;
		setFont("宋体", 12, false);
		setStyle(false, false);
	}

	/**
	 * 选择某行，如果该行存在则获取，不存在则创建
	 * 
	 * @param rownum
	 *            行号（从0开始计）
	 */
	private void selectRow(int rownum) {
		HSSFRow row = this.curWorkSheet.getRow(rownum);
		if (row == null) {
			row = this.curWorkSheet.createRow(rownum);
		}
		this.row = row; // 保存Row
		this.curRow = rownum; // 保存光标
	}

	/**
	 * 使用当前选择的Row,然后根据Column选择单元格
	 * 
	 * @param column
	 */
	private void selectCell(int column) {
		HSSFCell cell = this.row.getCell(column);
		if (cell == null) {
			cell = this.row.createCell(column);
		}
		this.cell = cell; // 保存Cell
		this.curColumn = column; // 保存光标
	}

	/**
	 * 获取当前光标的行值
	 * 
	 * @return
	 */
	public int getCurRow() {
		return curRow;
	}

	/**
	 * 获取列标题
	 * 
	 * @return
	 */
	public String[] getColsName() {
		return colsName;
	}

	/**
	 * 获取当前光标的列值
	 * 
	 * @return
	 */
	public int getCurColumn() {
		return curColumn;
	}

	/**
	 * 获取工作表数量
	 * 
	 * @return
	 */
	public int getNumberOfSheets() {
		return this.workbook.getNumberOfSheets();
	}

	/**
	 * 定位光标到指定的位置
	 * 
	 * @param row
	 *            行
	 * @param column
	 *            列
	 */
	public void locateCursor(int row, int column) {
		this.curRow = row;
		this.curColumn = column;
	}

	/**
	 * 在指定的单元格的位置写入值，写入后光标指向该单元格的右边单元格
	 * 
	 * @param rownum
	 *            行
	 * @param column
	 *            列
	 * @param value
	 *            值
	 */
	public void writeValue(int rownum, int column, Object value) {
		selectRow(rownum); // 选择行
		selectCell(column); // 选择列
		this.cell.setCellStyle(this.style);
		if (value instanceof Double) {
			this.cell.setCellValue((Double) value);
		} else {
			this.cell.setCellValue(value.toString());
		}
		this.curColumn++; // 将光标指向下一列
	}

	/**
	 * 在当前光标所在位置写入值， 写入后光标移动到当前单元格右边的单元格
	 * 
	 * @param value
	 *            值
	 */
	public void writeValue(Object value) {
		this.writeValue(this.curRow, this.curColumn, value);
	}

	/**
	 * 移动光标到下一行的第一个单元格的位置
	 */
	public void switchToNextRow() {
		this.curColumn = 0;
		this.curRow++;
	}

	/**
	 * 创建新的工作表，光标指向第一行第一列
	 */
	public void newSheet() {
		this.newSheet("工作表" + (this.workbook.getNumberOfSheets() + 1));
	}

	/**
	 * 创建新的工作表，光标指向第一行第一列
	 * 
	 * @param sheetName
	 *            工作表名称
	 */
	public void newSheet(String sheetName) {
		this.locateCursor(0, 0);
		this.curWorkSheet = this.workbook.createSheet(sheetName);
	}

	/**
	 * 切换当前的工作表，切换后光标指向第一行第一列
	 * 
	 * @param index
	 *            工作表序号
	 * @return
	 */
	public boolean switchSheet(int index) {
		this.locateCursor(0, 0);
		if (getNumberOfSheets() > index) {
			this.curWorkSheet = this.workbook.getSheetAt(index);
			return true;
		}
		return false;
	}

	public boolean generateExcelFile(String path) {
		FileOutputStream fileOut = null;
		try {
			fileOut = new FileOutputStream(path);
			workbook.write(fileOut);
		} catch (FileNotFoundException e) {
			e.printStackTrace();
			return false;
		} catch (IOException e) {
			e.printStackTrace();
			return false;
		} finally {
			try {
				if (fileOut != null)
					fileOut.close();
			} catch (IOException e) {
			}
		}
		return false;
	}

	/**
	 * 设置字体 <br/>
	 * 调用{@link #setStyle(boolean, boolean)}方法后生效
	 * 
	 * @param fontName
	 *            字体名称
	 * @param size
	 *            字体大小
	 * @param isBold
	 *            加粗
	 * @return
	 */
	public boolean setFont(String fontName, int size, boolean isBold) {
		return this.setFont(fontName, size, Font.COLOR_NORMAL, isBold, false,
				false, Font.U_NONE);
	}

	/**
	 * 设置字体 <br/>
	 * 调用{@link #setStyle(boolean, boolean)}方法后生效
	 * 
	 * @param fontName
	 *            字体名称
	 * @param size
	 *            字体大小
	 * @param color
	 *            字体颜色
	 * @param isBold
	 *            加粗
	 * @param isItalic
	 *            斜体
	 * @param strikeout
	 *            高亮
	 * @param underline
	 *            下划线
	 * @return
	 */
	public boolean setFont(String fontName, int size, int color,
			boolean isBold, boolean isItalic, boolean strikeout, byte underline) {
		HSSFFont font = workbook.createFont();
		font.setBold(isBold);
		font.setItalic(isItalic);
		font.setFontHeightInPoints((short) size);
		font.setColor((short) color);
		font.setStrikeout(strikeout);
		font.setUnderline(underline);
		this.font = font;

		return true;
	}

	/**
	 * 设置单元格样式（包括字体）
	 * 
	 * @param isCenter
	 *            是否居中显示
	 * @return
	 */
	public boolean setStyle(boolean isCenter, boolean withBorder) {
		HSSFCellStyle style = workbook.createCellStyle();
		if (withBorder) {
			style.setBorderBottom(BorderStyle.THIN);
			style.setBorderLeft(BorderStyle.THIN);
			style.setBorderRight(BorderStyle.THIN);
			style.setBorderTop(BorderStyle.MEDIUM_DASHED);
		}
		style.setFont(this.font);
		if (isCenter) {
			style.setAlignment(HorizontalAlignment.CENTER);
			style.setVerticalAlignment(VerticalAlignment.CENTER);
		}
		this.style = style;
		return true;
	}

	/**
	 * 设置某列为自动匹配宽度
	 * 
	 * @param column
	 */
	public void setAutoWidth(int column) {
		this.curWorkSheet.autoSizeColumn(column, true);
	}

	/**
	 * 合并单元格
	 * 
	 * @param value
	 *            值
	 * @param colSpan
	 *            占用列的宽度（单位：/列）
	 * @param rowSpan
	 *            占用行的宽度（单位：/行）
	 */
	public void writeAndMerge(String value, int colSpan, int rowSpan) {
		writeValue(value);
		this.curWorkSheet.addMergedRegion(new CellRangeAddress(curRow,
				curRow += --rowSpan, curColumn - 1, curColumn += --colSpan)); // 合并单元格
		this.curColumn++;
	}

	/**
	 * 向一行里面写入多个数据
	 * 
	 * @param switchNextLine
	 *            是否换行
	 * @param colsName
	 *            数据
	 */
	public void writeMultiValue(boolean switchNextLine, String... colsName) {
		if (switchNextLine) {
			switchToNextRow();
		}
		for (String var : colsName) {
			writeValue(var);
		}
	}

	/**
	 * 写入表格数据到Excel中
	 * 
	 * @param colsName
	 *            列的key
	 * @param content
	 *            Map集合，写入的内容
	 */
	public void writeTableContent(String[] colsName,
			List<Map<String, String>> content) {
		this.colsName = colsName;
		if (this.colsName == null || this.colsName.length == 0) {
			throw new RuntimeException("列名不能为空！[can't find the column's name!]"); // 列名不能为空
		}
		for (Map<String, String> rowValSet : content) {
			switchToNextRow();
			for (int i = 0; i < this.colsName.length; i++) {
				writeValue(rowValSet.get(this.colsName[i]));
			}
		}
	}

	@Override
	public void close() throws IOException {
		if (!isClosed)
			this.workbook.close();
	}
}
