package cc.cafetime;

import java.awt.Color;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.HashMap;
import java.util.Map;

import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.xssf.usermodel.DefaultIndexedColorMap;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFColor;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class ExcelWriter {

	public static final short EXCEL_COLUMN_WIDTH_FACTOR = 256;
	public static final int UNIT_OFFSET_LENGTH = 7;
	public static final int[] UNIT_OFFSET_MAP = new int[] { 0, 36, 73, 109, 146, 182, 219 };
	public static Map<String, XSSFCellStyle> map = new HashMap<>();

	public static void writeData(int[][][] data, String filePath) {
		try (XSSFWorkbook workbook = new XSSFWorkbook()) {
			XSSFSheet sheet = workbook.createSheet("student");
			int columnCount = data[0].length;
			// 遍历list,将数据写入Excel中
			for (int i = 0; i < data.length; i++) {
				if (i >= 250) {
					break;
				}
				long begin = System.currentTimeMillis();
				XSSFRow row = sheet.createRow(i);
				long end = System.currentTimeMillis();
				System.out.println("create row 第" + i + "行， 用时：" + (end - begin) + "ms");
				int[][] dataLine = data[i];
				for (int j = 0; j < dataLine.length; j++) {
					if (j >= 250) {
						break;
					}
					XSSFCell cell = row.createCell(j); // 写入一格
					int[] dataCell = dataLine[j];
					String colorStr = dataCell[0] + "," + dataCell[1] + "," + dataCell[2];
					XSSFCellStyle style = null;
					if (map.containsKey(colorStr)) {
						style = map.get(colorStr);
					} else {
						XSSFColor xssfColor = new XSSFColor(new Color(dataCell[0], dataCell[1], dataCell[2]),
								new DefaultIndexedColorMap());
						style = workbook.createCellStyle();
						long begin1 = System.currentTimeMillis();
						style.setFillForegroundColor(xssfColor);
						style.setFillPattern(FillPatternType.SOLID_FOREGROUND);
						end = System.currentTimeMillis();
						map.put(colorStr, style);
						System.out.println("fill color 第" + i + "行"+j+"列， 用时：" + (end - begin1) + "ms");
						
					}
					cell.setCellStyle(style);
				}
				end = System.currentTimeMillis();
				System.out.println("完成第" + i + "行， 用时：" + (end - begin) + "ms");
				row.setHeightInPoints(5);
			}
			for (int i = 0; i < columnCount; i++) {
				sheet.setColumnWidth(i, pixelWidth(5));
			}
			OutputStream out = null;
			out = new FileOutputStream(filePath);
			workbook.write(out);
			out.close();
		} catch (IOException e1) {
			e1.printStackTrace();
		}
	}

	/**
	 * 宽像素
	 * 
	 * @param pxs
	 * @return
	 */
	public static short pixelWidth(int pxs) {
		short width = (short) (EXCEL_COLUMN_WIDTH_FACTOR * (pxs / UNIT_OFFSET_LENGTH));
		width += UNIT_OFFSET_MAP[(pxs % UNIT_OFFSET_LENGTH)];
		return width;
	}

}
