package com.example.administrator.parseexcel;

import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;
import java.util.List;

import jxl.Cell;
import jxl.CellType;
import jxl.DateCell;
import jxl.NumberCell;
import jxl.Sheet;
import jxl.Workbook;

public class MainActivity extends AppCompatActivity {
    String tag = "wangweijun";
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);


    }


    public void readXLS(View view) {
        new Thread(new Runnable() {
            @Override
            public void run() {
                readXLS();
            }
        }).start();
    }

    public void readXlsx(View view) {
        new Thread(new Runnable() {
            @Override
            public void run() {
                readXlsx();
            }
        }).start();
    }


    private void readXLS() {
        try {
            Workbook workbook = null;
            try {
                // 注意一定要xls的扩展名
                File file=new File(Environment.getExternalStorageDirectory()+File.separator+"data2.xls");
                workbook = Workbook.getWorkbook(file);
//                InputStream inputStream= getAssets().open("data2.xls");
//                workbook=Workbook.getWorkbook(inputStream);
            } catch (Exception e) {
                e.printStackTrace();
                return;
            }
            //得到第一张表
            Sheet[] sheets = workbook.getSheets();
            int size = sheets.length;
            Log.i(tag, "sheets size:"+size);
            for (int i=0; i<size; i++) {
                Sheet sheet = workbook.getSheet(i);
                //列数
                int columnCount = sheet.getColumns();
                //行数
                int rowCount = sheet.getRows();
                //单元格
                Log.i(tag, "columnCount:"+columnCount+", rowCount:"+rowCount);
                Cell cell = null;
                for (int everyRow = 0; everyRow < rowCount; everyRow++) {
                    for (int everyColumn = 0; everyColumn < columnCount; everyColumn++) {
                        cell = sheet.getCell(everyColumn, everyRow);
                        if (cell.getType() == CellType.NUMBER) {
                            Log.i(tag, "数字="+ ((NumberCell) cell).getValue());
                        } else if (cell.getType() == CellType.DATE) {
                            Log.i(tag, "时间="+ ((DateCell) cell).getDate());
                        } else {
                            Log.i(tag, "everyColumn="+everyColumn+",everyRow="+everyRow+
                                    ",cell.getContents()="+ cell.getContents());
                        }
                    }
                }
            }

            //关闭workbook,防止内存泄露
            workbook.close();
        } catch (Exception e) {

        }
    }

    public void readXlsx() {
        InputStream stream = getResources().openRawResource(R.raw.data);
        try {
            XSSFWorkbook workbook = new XSSFWorkbook(stream);
            XSSFSheet sheet = workbook.getSheetAt(0);
            int rowsCount = sheet.getPhysicalNumberOfRows();
            Log.i("wangweijun", "rowsCount:"+rowsCount);
            FormulaEvaluator formulaEvaluator = workbook.getCreationHelper().createFormulaEvaluator();
            for (int r = 0; r<rowsCount; r++) {
                Row row = sheet.getRow(r);
                int cellsCount = row.getPhysicalNumberOfCells();
                for (int c = 0; c<cellsCount; c++) {
                    String value = getCellAsString(row, c, formulaEvaluator);
                    String cellInfo = "r:"+r+"; c:"+c+"; v:"+value;
                    Log.i("wangweijun", "value:"+value);
                }
            }
        } catch (Exception e) {
        }
    }

    protected String getCellAsString(Row row, int c, FormulaEvaluator formulaEvaluator) {
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
