package com.example.administrator.parseexcel;

import android.content.Intent;
import android.net.Uri;
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
                // 注意一定要xls的扩展名
				File file=new File(Environment.getExternalStorageDirectory()+File.separator+"data2.xls");
                JxlUtil.readXLS(file);
            }
        }).start();
    }

    public void readXlsx(View view) {
        new Thread(new Runnable() {
            @Override
            public void run() {
                File file=new File(Environment.getExternalStorageDirectory()+File.separator+"data.xlsx");
                PoiUtil.readXlsx(file);
            }
        }).start();
    }


    public void starTest(View view) {
//        startActivity(new Intent(getApplicationContext(), TestActivty.class));
        open();
    }

    public void open() {
        if (!Environment.getExternalStorageState().equals(Environment.MEDIA_MOUNTED)) {
            return;
        }
        //获取文件下载路径
        String path = Environment.getExternalStorageDirectory().getAbsolutePath() + "/";
        File dir = new File(path);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        //调用系统文件管理器打开指定路径目录
        Intent intent = new Intent(Intent.ACTION_GET_CONTENT);
        intent.setDataAndType(Uri.fromFile(dir.getParentFile()), "file/*.*");
        intent.addCategory(Intent.CATEGORY_OPENABLE);
        startActivityForResult(intent, REQUEST_CHOOSEFILE);

    }
   final int REQUEST_CHOOSEFILE = 0;
    @Override
    protected void onActivityResult(int requestCode,int resultCode,final Intent data){//选择文件返回
        super.onActivityResult(requestCode,resultCode,data);
        if(resultCode==RESULT_OK){
            switch(requestCode){
                case REQUEST_CHOOSEFILE:
                    new Thread(new Runnable() {
                        @Override
                        public void run() {
                            Uri uri=data.getData();
                            // /storage/emulated/0/data2.xls
                            Log.i(JxlUtil.TAG, "uri : "+uri.getPath()); // file:///storage/emulated/0/tvos.key
                            // /storage/emulated/0/data2.xls
                            File file = new File(uri.getPath());
                            String fileName = file.getName();
                            String suffixName = fileName.substring(fileName.lastIndexOf("."));
                            if (".xls".equals(suffixName)) {
                                JxlUtil.readXLS(file);
                            } else if (".xlsx".equals(suffixName)) {
                                PoiUtil.readXlsx(file);
                            } else {
                                Log.i(JxlUtil.TAG,"文件格式不支持");
                            }
                        }
                    }).start();
                    break;
            }
        }
    }

}
