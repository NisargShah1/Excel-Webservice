package org.ws.excel.service.impl;

import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Iterator;
import java.util.List;
import java.util.Map;

import jxl.Workbook;
import jxl.format.Alignment;
import jxl.format.Border;
import jxl.format.BorderLineStyle;
import jxl.format.Colour;
import jxl.write.Label;
import jxl.write.WritableCellFormat;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;

import org.apache.poi.hssf.usermodel.HSSFCellStyle;
import org.apache.poi.hssf.usermodel.HSSFFont;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.Font;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import org.springframework.stereotype.Service;
import org.ws.excel.model.Product;
import org.ws.excel.service.ifc.ExcelService;

@Service("excelService")
public class ExcelServiceImpl implements ExcelService{

	@Override
	public void createWorkbook(List recordsList,File file) {
		WritableWorkbook workBook = null;
		WritableSheet worksheet = null;
		WritableCellFormat cellFormatLock = new WritableCellFormat();
		WritableCellFormat cellFormatUnLock = new WritableCellFormat();
		Label cell = null;
		WritableCellFormat currCellFormat = null;
		try {
			workBook = Workbook.createWorkbook(file);
			worksheet = workBook.createSheet("Sheet1", 0);
			worksheet.setProtected(true);
			cellFormatLock.setLocked(true);
			cellFormatLock.setBackground(Colour.BLUE_GREY);
			cellFormatLock.setWrap(true);
			cellFormatLock.setBorder(Border.ALL, BorderLineStyle.THIN);
			cellFormatLock.setAlignment(Alignment.LEFT);
			cellFormatUnLock.setLocked(false);
			cellFormatUnLock.setWrap(true);
			cellFormatUnLock.setBorder(Border.ALL, BorderLineStyle.THIN);
			cellFormatUnLock.setAlignment(Alignment.LEFT);
			int rowi = 0;
			if (recordsList != null && !recordsList.isEmpty()) {
				Iterator itr = recordsList.iterator();
				Map recordsMap;
				boolean isHeader = true;
				while (itr.hasNext()) {
					recordsMap = (Map) itr.next();

					for (int iCount = 1; iCount <= recordsMap.size(); iCount++) {
						String cellValue = (recordsMap.get("COL" + iCount) != null ? recordsMap.get("COL" + iCount).toString() : "");


						if (isHeader) {
							currCellFormat = cellFormatLock;
						} else {
							currCellFormat = cellFormatUnLock;
						}

						cell = new Label(iCount - 1, rowi, cellValue.trim(), currCellFormat);
						worksheet.addCell(cell);

					}
					isHeader = false;
					rowi++;
				}
			}
			workBook.write();
			workBook.close();
		} catch (RowsExceededException ex) {
			ex.printStackTrace();
		} catch (WriteException ex) {
			ex.printStackTrace();
		} catch (IOException ex) {
			ex.printStackTrace();
		}

	}

	@Override
	public List<Map<String, String>> wriToExcel(List<Product> list) {
		List<Map<String,String>> result =  new ArrayList<Map<String,String>>();
		Map<String,String> map = new HashMap<String,String>();

		map.put("COL1", "ProductId");
		map.put("COL2", "Product Name");

		result.add(map);
		if (list.size() == 0) return result;
		for (int i = 0;i < list.size();i ++){
			Product product = list.get(i);
			map = new HashMap<String,String>();
			map.put("COL1", product.getProductId());
			map.put("COL2", product.getProductName());

			result.add(map);
		}
		return result;
	}

	@Override
	public void createWorkbookForPOI(List recordsList, File file) {
		HSSFWorkbook workBook = new HSSFWorkbook();
		Sheet sheet = workBook.createSheet("sheet1");
		CellStyle cellStyleLock = workBook.createCellStyle();
		HSSFFont font =workBook.createFont();
		font.setBoldweight(Font.BOLDWEIGHT_BOLD);
		cellStyleLock.setFillForegroundColor(HSSFColor.WHITE.index);
		cellStyleLock.setFillBackgroundColor(HSSFColor.BLUE.index);
		cellStyleLock.setFillPattern(HSSFCellStyle.SOLID_FOREGROUND);
		cellStyleLock.setFont(font);
		cellStyleLock.setLocked(true);
		CellStyle cellStyleUnLock = workBook.createCellStyle();
		cellStyleUnLock.setLocked(true);
		try {
			int rowi = 0;
			if (recordsList != null && !recordsList.isEmpty()) {
				Iterator itr = recordsList.iterator();
				Map recordsMap;
				boolean isHeader = true;
				while (itr.hasNext()) {
					recordsMap = (Map) itr.next();
					Row row = sheet.createRow(rowi);
					Cell cell =null;
					for (int iCount = 1; iCount <= recordsMap.size(); iCount++) {
						String cellValue = (recordsMap.get("COL" + iCount) != null ? recordsMap.get("COL" + iCount).toString() : "");

						cell = row.createCell(iCount - 1);
						cell.setCellValue(cellValue.trim());
						if (isHeader) {
							cell.setCellStyle(cellStyleLock);
						} else {
							cell.setCellStyle(cellStyleUnLock);
						}
					}
					isHeader = false;
					rowi++;
				}
			}

			FileOutputStream fileOut = null;

			fileOut = new FileOutputStream(file);

			workBook.write(fileOut);
			fileOut.flush();
			fileOut.close();
		} catch (FileNotFoundException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		} catch (IOException e) {
			// TODO Auto-generated catch block
			e.printStackTrace();
		}
	}


}
