package com.douwan.admin.util;

import android.os.Environment;
import android.util.Log;

import java.io.File;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;
public class JxlUtil {
	public static final String TAG = "wangweijun";
	public static  void readXLS(File file) {
		try {
			Workbook workbook = null;
			try {
				// 注意一定要xls的扩展名
//				File file=new File(Environment.getExternalStorageDirectory()+File.separator+"data2.xls");
				workbook = Workbook.getWorkbook(file);
//                InputStream inputStream= getAssets().open("data2.xls");
//                workbook=Workbook.getWorkbook(inputStream);
			} catch (Exception e) {
				e.printStackTrace();
				return;
			}
			Sheet[] sheets = workbook.getSheets();
			int size = sheets.length;
			Log.i(TAG, "sheets size:"+size);
			for (int i=0; i<size; i++) {
				Sheet sheet = workbook.getSheet(i);
				//列数
				int columnCount = sheet.getColumns();
				//行数
				int rowCount = sheet.getRows();
				//单元格
				Log.i(TAG, "columnCount:"+columnCount+", rowCount:"+rowCount);
				Cell cell = null;
				Log.i(TAG, "第 " + i + "表 ..."+sheet.getName());
				for (int everyRow = 0; everyRow < rowCount; everyRow++) {
					Log.i(TAG, "第 " + everyRow + "行");
					for (int everyColumn = 0; everyColumn < columnCount; everyColumn++) {
						cell = sheet.getCell(everyColumn, everyRow);
						if (cell.getType() == CellType.NUMBER) {
							Log.i(TAG, "数字="+ ((NumberCell) cell).getValue());
						} else if (cell.getType() == CellType.DATE) {
							Log.i(TAG, "时间="+ ((DateCell) cell).getDate());
						} else {
							Log.i(TAG, "everyColumn="+everyColumn+",everyRow="+everyRow+
									",cell.getContents()="+ cell.getContents());
						}
					}
				}
				Log.i(TAG, "第 " + i + "表 ... finished");
			}
			//关闭workbook,防止内存泄露
			workbook.close();
		} catch (Exception e) {

		}
	}
}
