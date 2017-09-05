package com.blackmcat.excel;

import java.io.File;
import java.io.FileOutputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.List;
import java.util.Map;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.util.CellRangeAddress;

import com.blackmcat.excelStyles.ExportExcelStyles;

public class ExportExcel {

	/**
	 * 导出excel的方法-xls
	 * 
	 * @param title
	 *            表格标题
	 * @param header
	 *            每一列的标题
	 * @param content
	 *            需要生成的值，List<List>;
	 */
	public static HSSFWorkbook getExcel_xls(String title, String[] header, List<List>... content) throws Exception {

		HSSFWorkbook workbook = new HSSFWorkbook(); // 生成Excel工作薄
		for (int cnum = 0; cnum < content.length; cnum++) {
			HSSFSheet sheet = workbook.createSheet(title+(cnum+1)); // 生成工作表
			Map<String, CellStyle> style = ExportExcelStyles.createStyles(workbook);// 导入样式
			sheet.addMergedRegion(CellRangeAddress.valueOf(ExcelUtil.mergeTitle(header.length))); // 合并单元格
			sheet.setDefaultColumnWidth(20); // 设置默认列宽度
			HSSFRow row0 = sheet.createRow(0);
			HSSFCell titleCell = row0.createCell(0);
			titleCell.setCellValue(title);
			titleCell.setCellStyle(style.get("title"));
			HSSFRow row1 = sheet.createRow(1);
			HSSFCell headerCell;
			for (int headerFlag = 0; headerFlag < header.length; headerFlag++) {
				headerCell = row1.createCell(headerFlag);
				headerCell.setCellValue(header[headerFlag]);
				headerCell.setCellStyle(style.get("header"));
			}

			HSSFRow row;
			HSSFCell cellCell;
			int rowFlag = 2;

			if (content != null && content[cnum].size() > 0) {
				for (int i = 0; i < content[cnum].size(); i++) {
					List li = (List) content[cnum].get(i);
					row = sheet.createRow(rowFlag);
					row.setHeightInPoints(20);
					for (int j = 0; j < li.size(); j++) {
						cellCell = row.createCell(j);
						cellCell.setCellValue(li.get(j).toString());
						cellCell.setCellType(HSSFCell.CELL_TYPE_STRING);
					}
					rowFlag++;
				}

			}
		}
		return workbook;
	}

	public static void main(String[] args) throws Exception {
		File f = new File("D://newExcel.xls");
		String[] title = { "1", "2", "3", "4", "5" };
		List<List> contens = new ArrayList<>();
		List<String> li = new ArrayList<>();
		li.add("11");
		li.add("2");
		li.add("3");
		li.add("4");
		li.add("5");
		contens.add(li);
		HSSFWorkbook xwb = getExcel_xls("newExcel", title, contens,contens);
		OutputStream out = new FileOutputStream(f);
		xwb.write(out);

	}
}
