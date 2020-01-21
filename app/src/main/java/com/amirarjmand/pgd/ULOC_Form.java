package com.amirarjmand.pgd;

import android.content.Context;
import android.content.Intent;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
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
import java.text.SimpleDateFormat;
import java.util.Date;
import java.util.Locale;

public class ULOC_Form extends AppCompatActivity {
    EditText c1,c2,c3,c4;
    Button save;
    String UnitNo;
    TextView header;
    String datea;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(ULOC_Form.this);
        setContentView(R.layout.activity_uloc__form);

        c1 = findViewById(R.id.C01);
        c2 = findViewById(R.id.C02);
        c3 = findViewById(R.id.C03);
        c4 = findViewById(R.id.C04);


        c4.setTextIsSelectable(true);
        
        header=findViewById(R.id.HeaderC) ;

        Bundle Uno=getIntent().getExtras();
        UnitNo=Uno.getString("uno");
        header.setText("Location information of unit  "+UnitNo);

        SimpleDateFormat sdf = new SimpleDateFormat("dd_MM_yyyy", Locale.getDefault());
        SimpleDateFormat simpleDateFormat=new SimpleDateFormat("dd.MMMM.yyyy");

//You can change "yyyyMMdd_HHmmss as per your requirement

        final String currentDateandTime = sdf.format(new Date());
        datea=simpleDateFormat.format(new Date());



        save=findViewById(R.id.Save);



        ReadDataFromExcel(ULOC_Form.this, "Units Location.xls");


        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveDatainxcel(ULOC_Form.this, "Units Location.xls");
                startActivity(new Intent(ULOC_Form.this, Unit_Loc_List.class));
                finish();

            }
        });

    }

    private void ReadDataFromExcel(ULOC_Form ULOC, String filename) {


        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Unit_Location/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook gov = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet GOV = gov.getSheetAt(0);

            if (UnitNo.equals("101")) {
                HSSFRow main = GOV.getRow(6);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("122")) {
                HSSFRow main = GOV.getRow(7);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("123")) {
                HSSFRow main = GOV.getRow(8);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("124")) {
                HSSFRow main = GOV.getRow(9);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("125")) {
                HSSFRow main = GOV.getRow(10);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("127")) {
                HSSFRow main = GOV.getRow(11);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("128")) {
                HSSFRow main = GOV.getRow(12);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("129")) {
                HSSFRow main = GOV.getRow(13);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("130")) {
                HSSFRow main = GOV.getRow(14);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("131")) {
                HSSFRow main = GOV.getRow(15);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("132")) {
                HSSFRow main = GOV.getRow(16);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("133")) {
                HSSFRow main = GOV.getRow(17);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("134")) {
                HSSFRow main = GOV.getRow(18);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("135")) {
                HSSFRow main = GOV.getRow(19);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

            if (UnitNo.equals("136")) {
                HSSFRow main = GOV.getRow(20);
                HSSFCell ConCell = main.getCell(1);
                HSSFCell ConCel2 = main.getCell(2);
                HSSFCell ConCel3 = main.getCell(3);
                HSSFCell ConCel4 = main.getCell(4);

                c1.setText(ConCell.toString());
                c2.setText(ConCel2.toString());
                c3.setText(ConCel3.toString());
                c4.setText(ConCel4.toString());
            }

        } catch (Exception e) {
            e.printStackTrace();
        }


    }
        private void SaveDatainxcel(Context context, String filename) {



            try {
                File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Unit_Location/", filename);
                FileInputStream myInputStream = new FileInputStream(file);

                HSSFWorkbook workbook = new HSSFWorkbook(myInputStream);



                HSSFSheet BUS_SYSTEM = workbook.getSheetAt(0);

                if (UnitNo.equals("101")){


                    HSSFRow cond = BUS_SYSTEM.getRow(6);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);


                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("122")){
                    HSSFRow cond = BUS_SYSTEM.getRow(7);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("123")){
                    HSSFRow cond = BUS_SYSTEM.getRow(8);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("124")){
                    HSSFRow cond = BUS_SYSTEM.getRow(9);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("125")){
                    HSSFRow cond = BUS_SYSTEM.getRow(10);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("127")){
                    HSSFRow cond = BUS_SYSTEM.getRow(11);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("128")){
                    HSSFRow cond = BUS_SYSTEM.getRow(12);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("129")){
                    HSSFRow cond = BUS_SYSTEM.getRow(13);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("130")){
                    HSSFRow cond = BUS_SYSTEM.getRow(14);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("131")){
                    HSSFRow cond = BUS_SYSTEM.getRow(15);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("132")){
                    HSSFRow cond = BUS_SYSTEM.getRow(16);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("133")){
                    HSSFRow cond = BUS_SYSTEM.getRow(17);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("134")){
                    HSSFRow cond = BUS_SYSTEM.getRow(18);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("135")){
                    HSSFRow cond = BUS_SYSTEM.getRow(19);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());
                }

                if (UnitNo.equals("136")){
                    HSSFRow cond = BUS_SYSTEM.getRow(20);
                    HSSFCell con1 = cond.getCell(1);
                    HSSFCell con2 = cond.getCell(2);
                    HSSFCell con3 = cond.getCell(3);
                    HSSFCell con4 = cond.getCell(4);

                    con1.setCellValue(c1.getText().toString().trim());
                    con2.setCellValue(c2.getText().toString().trim());
                    con3.setCellValue(c3.getText().toString().trim());
                    con4.setCellValue(c4.getText().toString().trim());   }

                    myInputStream.close();
                FileOutputStream fos = new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Unit_Location/", filename));
                workbook.write(fos);
                fos.close();
                Toast.makeText(context, "Changes are done.", Toast.LENGTH_SHORT).show();
            } catch (Exception e) {
                e.printStackTrace();
            }
        }

    }



