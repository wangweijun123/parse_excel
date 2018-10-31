package com.example.administrator.parseexcel;

import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.Log;
import android.view.View;

import com.douwan.admin.util.JxlUtil;
import com.douwan.admin.util.PoiUtil;

import org.apache.poi.hssf.usermodel.HSSFDateUtil;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.InputStream;
import java.text.SimpleDateFormat;

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
                JxlUtil.readXLS();
            }
        }).start();
    }

    public void readXlsx(View view) {
        new Thread(new Runnable() {
            @Override
            public void run() {
                PoiUtil.readXlsx();
            }
        }).start();
    }






}
