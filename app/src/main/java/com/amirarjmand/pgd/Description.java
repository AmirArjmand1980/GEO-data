package com.amirarjmand.pgd;

import android.content.Context;
import android.support.annotation.NonNull;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.CalendarView;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Description extends AppCompatActivity {
    EditText c2;

    Button save;
    String UnitNo, date;
    TextView header;
    CalendarView calendarView;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(Description.this);
        setContentView(R.layout.activity_description);


        c2=findViewById(R.id.Co2);
        calendarView = findViewById(R.id.Calen);

        calendarView.setOnDateChangeListener(new CalendarView.OnDateChangeListener() {
            @Override
            public void onSelectedDayChange(@NonNull CalendarView calendarView, int year, int month, int dayOfMonth) {
                date = dayOfMonth + "." + (month + 1) + "." + year;
            }
        });

        header=findViewById(R.id.HeaderC) ;

        Bundle Uno=getIntent().getExtras();
        UnitNo=Uno.getString("uno");
        header.setText("Sending Information of Unit  "+UnitNo);


        save=findViewById(R.id.Save);


        ReadDataFromExcel(Description.this, "Inventory List unit "+UnitNo+".xls");

        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveDatainxcel(Description.this, "Inventory List unit "+UnitNo+".xls");

            }
        });

    }
    private void ReadDataFromExcel(Description Description, String filename) {

        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook inv = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet BUS_SYSTEM = inv.getSheetAt(0);




            HSSFRow Eva = BUS_SYSTEM.getRow(9);
            HSSFCell eval = Eva.getCell(2);
            c2.setText(eval.toString());

        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void SaveDatainxcel(Context context, String filename) {

        try {
            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);

            HSSFWorkbook workbook = new HSSFWorkbook (myInputStream);
            HSSFSheet CONTENT = workbook.getSheetAt(0);

            HSSFRow cond = CONTENT.getRow(8);
            HSSFCell con3 = cond.getCell(2);
            con3.setCellValue("Unit sent to "+c2.getText().toString().trim()+"at "+date);

            


            myInputStream.close();
            FileOutputStream fos =new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename));
            workbook.write(fos);
            fos.close();
            Toast.makeText(context, "Description Saved", Toast.LENGTH_SHORT).show();
        }
        catch (Exception e){e.printStackTrace();}
    }
}
