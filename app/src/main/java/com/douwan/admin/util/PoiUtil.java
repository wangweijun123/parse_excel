package com.douwan.admin.util;

import android.os.Environment;
import android.util.Log;
import java.io.File;
import java.text.SimpleDateFormat;
import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

public class PoiUtil {
	public static final String TAG = "wangweijun";
	public static void readXlsx() {
//        InputStream stream = getResources().openRawResource(R.raw.data);
//        XSSFWorkbook workbook = new XSSFWorkbook(stream);
		try {
//            data.xlsx
			File file=new File(Environment.getExternalStorageDirectory()+File.separator+"data.xlsx");
			XSSFWorkbook workbook = new XSSFWorkbook(file);
			int sheets = workbook.getNumberOfSheets();
			Log.i(TAG, "一共有表 : "+sheets);
			for (int i=0; i<sheets; i++) {

				XSSFSheet sheet = workbook.getSheetAt(i);
                Log.i(TAG, "表開始 ..."+i + " " +  sheet.getSheetName());
				int rowsCount = sheet.getPhysicalNumberOfRows();
				Log.i(TAG, "rowsCount:"+rowsCount);
				FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
				for (int r = 0; r<rowsCount; r++) {
					Row row = sheet.getRow(r);
					int cellsCount = row.getPhysicalNumberOfCells();
					for (int c = 0; c<cellsCount; c++) {
						String value = getCellAsString(row, c, formulaEvaluator);
						String cellInfo = "r:"+r+"; c:"+c+"; v:"+value;
						Log.i(TAG, cellInfo);
					}
				}

				Log.i(TAG, "表finished ..."+i);
			}
		} catch (Exception e) {
		}
	}

	private static String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
		String value = "";
		try {
			org.apache.poi.ss.usermodel.Cell cell = row.getCell(c);
			CellValue cellValue = formulaEvaluator.evaluate(cell);
			switch (cellValue.getCellType()) {
				case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_BOOLEAN:
					value = ""+cellValue.getBooleanValue();
					break;
				case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_NUMERIC:
					double numericValue = cellValue.getNumberValue();
					if(HSSFDateUtil.isCellDateFormatted(cell)) {
						double date = cellValue.getNumberValue();
						SimpleDateFormat formatter =
								new SimpleDateFormat("dd/MM/yy");
						value = formatter.format(HSSFDateUtil.getJavaDate(date));
					} else {
						value = ""+numericValue;
					}
					break;
				case org.apache.poi.ss.usermodel.Cell.CELL_TYPE_STRING:
					value = ""+cellValue.getStringValue();
					break;
				default:
			}
		} catch (NullPointerException e) {
			/* proper error handling should be here */
		}
		return value;
	}
}
