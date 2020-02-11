package com.amirarjmand.pgd.Inventory;

import android.content.Context;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;
import android.widget.Toast;

import com.amirarjmand.pgd.R;
import com.amirarjmand.pgd.noTitle;

import org.apache.poi.hssf.usermodel.HSSFCell;
import org.apache.poi.hssf.usermodel.HSSFRow;
import org.apache.poi.hssf.usermodel.HSSFSheet;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.poifs.filesystem.POIFSFileSystem;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;

public class Alarm extends AppCompatActivity {

    EditText c1,c2,c3,c4,c5,c6,c7,c8,c9,c10,c11,c12,c13,c14,c15,c16,c17,c18,c19,c20,c21,c22,c23,c24,c25,c26,c27,c28,c29,c30,c31,c32,c33,c34,c35,c36,c37,c38,c39,c40,c41,c42,c43,c44,c45,c46,c47,c48,c49,c50,c51,c52,c53,c54;

    Button save;
    String UnitNo, filename;
    TextView header;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(Alarm.this);
        setContentView(R.layout.activity_alarm);

        c1=findViewById(R.id.Co1);
        c2=findViewById(R.id.Co2);
        c3=findViewById(R.id.Co3);
        c4=findViewById(R.id.Co4);
        c5=findViewById(R.id.Co5);
        c6=findViewById(R.id.Co6);
        c7=findViewById(R.id.Co7);
        c8=findViewById(R.id.Co8);
        c9=findViewById(R.id.Co9);
        c10=findViewById(R.id.Co10);
        c11=findViewById(R.id.Co11);
        c12=findViewById(R.id.Co12);
        c13=findViewById(R.id.Co13);
        c14=findViewById(R.id.Co14);
        c15=findViewById(R.id.Co15);
        c16=findViewById(R.id.Co16);
        c17=findViewById(R.id.Co17);
        c18=findViewById(R.id.Co18);
        c19=findViewById(R.id.Co19);
        c20=findViewById(R.id.Co20);
        c21=findViewById(R.id.Co21);
        c22=findViewById(R.id.Co22);
        c23=findViewById(R.id.Co23);
        c24=findViewById(R.id.Co24);
        c25=findViewById(R.id.Co25);
        c26=findViewById(R.id.Co26);
        c27=findViewById(R.id.Co27);
        c28=findViewById(R.id.Co28);
        c29=findViewById(R.id.Co29);
        c30=findViewById(R.id.Co30);
        c31=findViewById(R.id.Co31);
        c32=findViewById(R.id.Co32);
        c33=findViewById(R.id.Co33);
        c34=findViewById(R.id.Co34);
        c35=findViewById(R.id.Co35);
        c36=findViewById(R.id.Co36);
        c37=findViewById(R.id.Co37);
        c38=findViewById(R.id.Co38);
        c39=findViewById(R.id.Co39);
        c40=findViewById(R.id.Co40);
        c41=findViewById(R.id.Co41);
        c42=findViewById(R.id.Co42);
        c43=findViewById(R.id.Co43);
        c44=findViewById(R.id.Co44);
        c45=findViewById(R.id.Co45);
        c46=findViewById(R.id.Co46);
        c47=findViewById(R.id.Co47);
        c48=findViewById(R.id.Co48);
        c49=findViewById(R.id.Co49);
        c50=findViewById(R.id.Co50);
        c51=findViewById(R.id.Co51);
        c52=findViewById(R.id.Co52);
        c53=findViewById(R.id.Co53);
        c54=findViewById(R.id.Co54);


        header=findViewById(R.id.HeaderC) ;

        Bundle Uno=getIntent().getExtras();
        UnitNo=Uno.getString("uno");
        header.setText("Alarm system of unit "+UnitNo);

        filename=  "Inventory List unit"+UnitNo+".xls";

        save=findViewById(R.id.Save);

        ReadDataFromExcel(Alarm.this, "Inventory List unit "+UnitNo+".xls");

        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                SaveDatainxcel(Alarm.this, "Inventory List unit "+UnitNo+".xls");

            }
        });



    }
    private void ReadDataFromExcel(Alarm Alarm, String filename) {

        try {


            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook inv = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet BUS_SYSTEM = inv.getSheetAt(5);

            HSSFRow Condenser = BUS_SYSTEM.getRow(5);
            HSSFCell ConCell = Condenser.getCell(4);
            HSSFCell ConCel2 = Condenser.getCell(5);
            HSSFCell ConCel3 = Condenser.getCell(3);
            c1.setText(ConCell.toString());
            c2.setText(ConCel2.toString());
            c3.setText(ConCel3.toString());

            HSSFRow Eva = BUS_SYSTEM.getRow(6);
            HSSFCell eval = Eva.getCell(4);
            HSSFCell eva2 = Eva.getCell(5);
            HSSFCell eva3 = Eva.getCell(3);
            c4.setText(eval.toString());
            c5.setText(eva2.toString());
            c6.setText(eva3.toString());

            HSSFRow comp = BUS_SYSTEM.getRow(7);
            HSSFCell com1 = comp.getCell(4);
            HSSFCell com2 = comp.getCell(5);
            HSSFCell com3 = comp.getCell(3);
            c7.setText(com1.toString());
            c8.setText(com2.toString());
            c9.setText(com3.toString());

            HSSFRow mult = BUS_SYSTEM.getRow(8);
            HSSFCell mul1 = mult.getCell(4);
            HSSFCell mul2 = mult.getCell(5);
            HSSFCell mul3 = mult.getCell(3);
            c10.setText(mul1.toString());
            c11.setText(mul2.toString());
            c12.setText(mul3.toString());

            HSSFRow ligh = BUS_SYSTEM.getRow(9);
            HSSFCell lig1 = ligh.getCell(4);
            HSSFCell lig2 = ligh.getCell(5);
            HSSFCell lig3 = ligh.getCell(3);
            c13.setText(lig1.toString());
            c14.setText(lig2.toString());
            c15.setText(lig3.toString());

            HSSFRow term = BUS_SYSTEM.getRow(10);
            HSSFCell ter1 = term.getCell(4);
            HSSFCell ter2 = term.getCell(5);
            HSSFCell ter3 = term.getCell(3);
            c16.setText(ter1.toString());
            c17.setText(ter2.toString());
            c18.setText(ter3.toString());

            HSSFRow fusi = BUS_SYSTEM.getRow(11);
            HSSFCell fus1 = fusi.getCell(4);
            HSSFCell fus2 = fusi.getCell(5);
            HSSFCell fus3 = fusi.getCell(3);
            c19.setText(fus1.toString());
            c20.setText(fus2.toString());
            c21.setText(fus3.toString());

            HSSFRow powr = BUS_SYSTEM.getRow(12);
            HSSFCell pow1 = powr.getCell(4);
            HSSFCell pow2 = powr.getCell(5);
            HSSFCell pow3 = powr.getCell(3);
            c22.setText(pow1.toString());
            c23.setText(pow2.toString());
            c24.setText(pow3.toString());

            HSSFRow trns = BUS_SYSTEM.getRow(13);
            HSSFCell trn1 = trns.getCell(4);
            HSSFCell trn2 = trns.getCell(5);
            HSSFCell trn3 = trns.getCell(3);
            c25.setText(trn1.toString());
            c26.setText(trn2.toString());
            c27.setText(trn3.toString());

            HSSFRow fani = BUS_SYSTEM.getRow(14);
            HSSFCell fan1 = fani.getCell(4);
            HSSFCell fan2 = fani.getCell(5);
            HSSFCell fan3 = fani.getCell(3);
            c28.setText(fan1.toString());
            c29.setText(fan2.toString());
            c30.setText(fan3.toString());

            HSSFRow upsi = BUS_SYSTEM.getRow(15);
            HSSFCell ups1 = upsi.getCell(4);
            HSSFCell ups2 = upsi.getCell(5);
            HSSFCell ups3 = upsi.getCell(3);
            c31.setText(ups1.toString());
            c32.setText(ups2.toString());
            c33.setText(ups3.toString());

            HSSFRow purg = BUS_SYSTEM.getRow(16);
            HSSFCell pur1 = purg.getCell(4);
            HSSFCell pur2 = purg.getCell(5);
            HSSFCell pur3 = purg.getCell(3);
            c34.setText(pur1.toString());
            c35.setText(pur2.toString());
            c36.setText(pur3.toString());

            HSSFRow smok = BUS_SYSTEM.getRow(17);
            HSSFCell smo1 = smok.getCell(4);
            HSSFCell smo2 = smok.getCell(5);
            HSSFCell smo3 = smok.getCell(3);
            c37.setText(smo1.toString());
            c38.setText(smo2.toString());
            c39.setText(smo3.toString());

            HSSFRow fire = BUS_SYSTEM.getRow(18);
            HSSFCell fir1 = fire.getCell(4);
            HSSFCell fir2 = fire.getCell(5);
            HSSFCell fir3 = fire.getCell(3);
            c40.setText(fir1.toString());
            c41.setText(fir2.toString());
            c42.setText(fir3.toString());

            HSSFRow toxi = BUS_SYSTEM.getRow(19);
            HSSFCell tox1 = toxi.getCell(4);
            HSSFCell tox2 = toxi.getCell(5);
            HSSFCell tox3 = toxi.getCell(3);
            c43.setText(tox1.toString());
            c44.setText(tox2.toString());
            c45.setText(tox3.toString());

            HSSFRow flam = BUS_SYSTEM.getRow(20);
            HSSFCell flm1 = flam.getCell(4);
            HSSFCell flm2 = flam.getCell(5);
            HSSFCell flm3 = flam.getCell(3);
            c46.setText(flm1.toString());
            c47.setText(flm2.toString());
            c48.setText(flm3.toString());

            HSSFRow firs = BUS_SYSTEM.getRow(21);
            HSSFCell fira1 = firs.getCell(4);
            HSSFCell fira2 = firs.getCell(5);
            HSSFCell fira3 = firs.getCell(3);
            c49.setText(fira1.toString());
            c50.setText(fira2.toString());
            c51.setText(fira3.toString());

            HSSFRow radi = BUS_SYSTEM.getRow(22);
            HSSFCell rad1 = radi.getCell(4);
            HSSFCell rad2 = radi.getCell(5);
            HSSFCell rad3 = radi.getCell(3);
            c52.setText(rad1.toString());
            c53.setText(rad2.toString());
            c54.setText(rad3.toString());

        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void SaveDatainxcel(Context context, String filename) {

        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);

            HSSFWorkbook workbook = new HSSFWorkbook(myInputStream);
            HSSFSheet BUS_SYSTEM = workbook.getSheetAt(5);

            HSSFRow cond = BUS_SYSTEM.getRow(5);
            HSSFCell con1 = cond.getCell(4);
            HSSFCell con2 = cond.getCell(5);
            HSSFCell con3 = cond.getCell(3);
            con1.setCellValue(c1.getText().toString().trim());
            con2.setCellValue(c2.getText().toString().trim());
            con3.setCellValue(c3.getText().toString().trim());

            HSSFRow evap = BUS_SYSTEM.getRow(6);
            HSSFCell eva1 = evap.getCell(4);
            HSSFCell eva2 = evap.getCell(5);
            HSSFCell eva3 = evap.getCell(3);
            eva1.setCellValue(c4.getText().toString().trim());
            eva2.setCellValue(c5.getText().toString().trim());
            eva3.setCellValue(c6.getText().toString().trim());

            HSSFRow comp = BUS_SYSTEM.getRow(7);
            HSSFCell com1 = comp.getCell(4);
            HSSFCell com2 = comp.getCell(5);
            HSSFCell com3 = comp.getCell(3);
            com1.setCellValue(c7.getText().toString().trim());
            com2.setCellValue(c8.getText().toString().trim());
            com3.setCellValue(c9.getText().toString().trim());

            HSSFRow mult = BUS_SYSTEM.getRow(8);
            HSSFCell mul1 = mult.getCell(4);
            HSSFCell mul2 = mult.getCell(5);
            HSSFCell mul3 = mult.getCell(3);
            mul1.setCellValue(c10.getText().toString().trim());
            mul2.setCellValue(c11.getText().toString().trim());
            mul3.setCellValue(c12.getText().toString().trim());

            HSSFRow ligt = BUS_SYSTEM.getRow(9);
            HSSFCell lig1 = ligt.getCell(4);
            HSSFCell lig2 = ligt.getCell(5);
            HSSFCell lig3 = ligt.getCell(3);
            lig1.setCellValue(c13.getText().toString().trim());
            lig2.setCellValue(c14.getText().toString().trim());
            lig3.setCellValue(c15.getText().toString().trim());

            HSSFRow term = BUS_SYSTEM.getRow(10);
            HSSFCell ter1 = term.getCell(4);
            HSSFCell ter2 = term.getCell(5);
            HSSFCell ter3 = term.getCell(3);
            ter1.setCellValue(c16.getText().toString().trim());
            ter2.setCellValue(c17.getText().toString().trim());
            ter3.setCellValue(c18.getText().toString().trim());

            HSSFRow fusi = BUS_SYSTEM.getRow(11);
            HSSFCell fus1 = fusi.getCell(4);
            HSSFCell fus2 = fusi.getCell(5);
            HSSFCell fus3 = fusi.getCell(3);
            fus1.setCellValue(c19.getText().toString().trim());
            fus2.setCellValue(c20.getText().toString().trim());
            fus3.setCellValue(c21.getText().toString().trim());

            HSSFRow freq = BUS_SYSTEM.getRow(12);
            HSSFCell fre1 = freq.getCell(4);
            HSSFCell fre2 = freq.getCell(5);
            HSSFCell fre3 = freq.getCell(3);
            fre1.setCellValue(c22.getText().toString().trim());
            fre2.setCellValue(c23.getText().toString().trim());
            fre3.setCellValue(c24.getText().toString().trim());

            HSSFRow powr = BUS_SYSTEM.getRow(13);
            HSSFCell pow1 = powr.getCell(4);
            HSSFCell pow2 = powr.getCell(5);
            HSSFCell pow3 = powr.getCell(3);
            pow1.setCellValue(c25.getText().toString().trim());
            pow2.setCellValue(c26.getText().toString().trim());
            pow3.setCellValue(c27.getText().toString().trim());

            HSSFRow ufan = BUS_SYSTEM.getRow(14);
            HSSFCell fan1 = ufan.getCell(4);
            HSSFCell fan2 = ufan.getCell(5);
            HSSFCell fan3 = ufan.getCell(3);
            fan1.setCellValue(c28.getText().toString().trim());
            fan2.setCellValue(c29.getText().toString().trim());
            fan3.setCellValue(c30.getText().toString().trim());

            HSSFRow upsi = BUS_SYSTEM.getRow(15);
            HSSFCell ups1 = upsi.getCell(4);
            HSSFCell ups2 = upsi.getCell(5);
            HSSFCell ups3 = upsi.getCell(3);
            ups1.setCellValue(c31.getText().toString().trim());
            ups2.setCellValue(c32.getText().toString().trim());
            ups3.setCellValue(c33.getText().toString().trim());

            HSSFRow purg = BUS_SYSTEM.getRow(16);
            HSSFCell pur1 = purg.getCell(4);
            HSSFCell pur2 = purg.getCell(5);
            HSSFCell pur3 = purg.getCell(3);
            pur1.setCellValue(c34.getText().toString().trim());
            pur2.setCellValue(c35.getText().toString().trim());
            pur3.setCellValue(c36.getText().toString().trim());

            HSSFRow smok = BUS_SYSTEM.getRow(17);
            HSSFCell smk1 = smok.getCell(4);
            HSSFCell smk2 = smok.getCell(5);
            HSSFCell smk3 = smok.getCell(3);
            smk1.setCellValue(c37.getText().toString().trim());
            smk2.setCellValue(c38.getText().toString().trim());
            smk3.setCellValue(c39.getText().toString().trim());

            HSSFRow fire = BUS_SYSTEM.getRow(18);
            HSSFCell fire1 = fire.getCell(4);
            HSSFCell fire2 = fire.getCell(5);
            HSSFCell fire3 = fire.getCell(3);
            fire1.setCellValue(c40.getText().toString().trim());
            fire2.setCellValue(c41.getText().toString().trim());
            fire3.setCellValue(c42.getText().toString().trim());

            HSSFRow toxi = BUS_SYSTEM.getRow(19);
            HSSFCell tox1 = toxi.getCell(4);
            HSSFCell tox2 = toxi.getCell(5);
            HSSFCell tox3 = toxi.getCell(3);
            tox1.setCellValue(c43.getText().toString().trim());
            tox2.setCellValue(c44.getText().toString().trim());
            tox3.setCellValue(c45.getText().toString().trim());

            HSSFRow flam = BUS_SYSTEM.getRow(20);
            HSSFCell flm1 = flam.getCell(4);
            HSSFCell flm2 = flam.getCell(5);
            HSSFCell flm3 = flam.getCell(3);
            flm1.setCellValue(c46.getText().toString().trim());
            flm2.setCellValue(c47.getText().toString().trim());
            flm3.setCellValue(c48.getText().toString().trim());

            HSSFRow faid = BUS_SYSTEM.getRow(21);
            HSSFCell fai1 = faid.getCell(4);
            HSSFCell fai2 = faid.getCell(5);
            HSSFCell fai3 = faid.getCell(3);
            fai1.setCellValue(c49.getText().toString().trim());
            fai2.setCellValue(c50.getText().toString().trim());
            fai3.setCellValue(c51.getText().toString().trim());

            HSSFRow radi = BUS_SYSTEM.getRow(22);
            HSSFCell rad1 = radi.getCell(4);
            HSSFCell rad2 = radi.getCell(5);
            HSSFCell rad3 = radi.getCell(3);
            rad1.setCellValue(c52.getText().toString().trim());
            rad2.setCellValue(c53.getText().toString().trim());
            rad3.setCellValue(c54.getText().toString().trim());

            myInputStream.close();
            FileOutputStream fos = new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename));
            workbook.write(fos);
            fos.close();
            Toast.makeText(context, "Changes are done.", Toast.LENGTH_SHORT).show();
        }
        catch (Exception e){e.printStackTrace();}
    }
}
