package com.administrator.jxl;


import android.content.DialogInterface;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.Switch;
import android.widget.TextView;
import android.widget.Toast;


import java.io.File;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.util.Random;

import jxl.Cell;
import jxl.Sheet;
import jxl.Workbook;
import jxl.read.biff.BiffException;
import jxl.write.Label;
import jxl.write.WritableSheet;
import jxl.write.WritableWorkbook;
import jxl.write.WriteException;
import jxl.write.biff.RowsExceededException;


public class MainActivity extends AppCompatActivity implements View.OnClickListener {
    String[] str = new String[]{"one", "two", "three", "four", "five", "six", "seven", "eight", "nine", "ten"};
    File file;
    Person person;
    Button add;
    Button ser;
    Button del;
    TextView text;
    int num=0;
    @Override

    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        ;
        text=(TextView)findViewById(R.id.excel);
        add=(Button)findViewById(R.id.adda);
        del=(Button)findViewById(R.id.delate);
        ser=(Button)findViewById(R.id.search);
        add.setOnClickListener(this);
        del.setOnClickListener(this);
        ser.setOnClickListener(this);
        //person=new Person(12,"Mr wang","男");
        file=new File(Environment.getExternalStorageDirectory().getPath()+"/ExcelTest");
        if (!file.exists()) {
            file.mkdir();
        }
       file = new File(Environment.getExternalStorageDirectory().getPath() + "/ExcelTest/ExcelTest.xls");//创建文件夹
        if(!file.exists()){
            createFile(file);
            ;
            }
    }

public void createFile(File file) {
    WritableWorkbook wookbook = null;
    try {
        file.createNewFile();
        //没必要用文件流
        wookbook = Workbook.createWorkbook(file);
        WritableSheet sheet = wookbook.createSheet("第一张表", 0);
        Label lable1 = new Label(0, 0, "姓名");
        Label lable2 = new Label(1, 0, "年龄");
        Label lable3 = new Label(2, 0, "性别");
        sheet.addCell(lable1);
        sheet.addCell(lable2);
        sheet.addCell(lable3);
        Cell cell = sheet.getWritableCell(0, 0);
        wookbook.write();
        wookbook.close();
    } catch (Exception e) {
        e.printStackTrace();
    } finally {

    }

}
     public void addPerson(Person person,File file) {
         Workbook original=null;
         WritableWorkbook workbook=null;
         try {//  如果是想要修改一个已存在的excel工作簿，则需要先获得它的原始工作簿，再创建一个可读写的副本：
             original=Workbook.getWorkbook(file);
             workbook=Workbook.createWorkbook(file,original);
             WritableSheet sheet=workbook.getSheet(0);
             int row=sheet.getRows();
             Label label=new Label(0,row,person.getName());
             Label labe2=new Label(1,row,person.getAge());
             Label labe3=new Label(2,row,person.getSex());
             sheet.addCell(label);
             sheet.addCell(labe2);
             sheet.addCell(labe3);
             workbook.write();

         }catch (Exception e){

         }finally {
             if(original!=null){
                 original.close();
             }
             if(workbook!=null){
                 try {
                     workbook.close();
                 } catch (IOException e) {
                     e.printStackTrace();
                 } catch (WriteException e) {
                     e.printStackTrace();
                 }
             }

         }

     }

     public void  readPerson(File file){

        Workbook wookbook=null;
        try {
            wookbook=Workbook.getWorkbook(file);
             Sheet sheet=wookbook.getSheet(0);//只读，可不建立副本
             int i=sheet.getRows();
             if(num<=(i-1)) {

                 Cell cell = sheet.getCell(1, num);
                 text.setText(cell.getContents());
                 num++;
             }
             else{
                 num=0;
             }

        } catch (IOException e) {
            e.printStackTrace();
        } catch (BiffException e) {
            e.printStackTrace();
        }finally {
            wookbook.close();
        }


     }

     public  void delatePerson(File file){
        Workbook originalbook=null;
        WritableWorkbook newbook=null;
        try {
            originalbook=Workbook.getWorkbook(file);
            newbook=Workbook.createWorkbook(file,originalbook);
            WritableSheet sheet=newbook.getSheet(0);
            int row=sheet.getRows();
            if(row>1)
              sheet.removeRow(row-1);
            newbook.write();
        }catch (Exception e){


        }
        finally {
            if(originalbook!=null){
                originalbook.close();
            }
            if(newbook!=null) {
                try {
                    newbook.close();
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }

        }


     }
    @Override
    public void onClick(View v) {
        switch (v.getId()){
            case R.id.adda:
                Person person = new Person(new Random().nextInt(20)+"", "王显彬",  "男");
                addPerson(person,file);
                    break;
            case R.id.delate:
                delatePerson(file);

                break;
            case R.id.search:
                readPerson(file);

                break;


        }
    }
}

