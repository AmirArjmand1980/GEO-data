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

public class Shortage extends AppCompatActivity {
    EditText e1, e2, e3, e4, e5, e6, e7, e8, e9, e10, e11, e12, e13, e14, e15, e16, e17, e18, e19, e20, e21, e22, e23, e24, e25, e26, e27, e28, e29, e30, e31, e32, e33, e34, e35, e36, e37, e38, e39, e40, e41, e42, e43, e44, e45, e46, e47, e48, e49, e50, e51, e52, e53, e54, e55, e56, e57, e58, e59, e60, e61, e62, e63, e64, e65, e66, e67, e68, e69, e70, e71, e72, e73, e74, e75, e76, e77, e78, e79, e80, e81, e82, e83, e84, e85, e86, e87, e88, e89, e90, e91, e92, e93;


    Button save;
    String UnitNo;
    TextView header;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(Shortage.this);
        setContentView(R.layout.activity_shortage);

        e1 = findViewById(R.id.E1);
        e2 = findViewById(R.id.E2);
        e3 = findViewById(R.id.E3);
        e4 = findViewById(R.id.E4);
        e5 = findViewById(R.id.E5);
        e6 = findViewById(R.id.E6);
        e7 = findViewById(R.id.E7);
        e8 = findViewById(R.id.E8);
        e9 = findViewById(R.id.E9);
        e10 = findViewById(R.id.E10);
        e11 = findViewById(R.id.E11);
        e12 = findViewById(R.id.E12);
        e13 = findViewById(R.id.E13);
        e14 = findViewById(R.id.E14);
        e15 = findViewById(R.id.E15);
        e16 = findViewById(R.id.E16);
        e17 = findViewById(R.id.E17);
        e18 = findViewById(R.id.E18);
        e19 = findViewById(R.id.E19);
        e20 = findViewById(R.id.E20);
        e21 = findViewById(R.id.E21);
        e22 = findViewById(R.id.E22);
        e23 = findViewById(R.id.E23);
        e24 = findViewById(R.id.E24);
        e25 = findViewById(R.id.E25);
        e26 = findViewById(R.id.E26);
        e27 = findViewById(R.id.E27);
        e28 = findViewById(R.id.E28);
        e29 = findViewById(R.id.E29);
        e30 = findViewById(R.id.E30);
        e31 = findViewById(R.id.E31);
        e32 = findViewById(R.id.E32);
        e33 = findViewById(R.id.E33);
        e34 = findViewById(R.id.E34);
        e35 = findViewById(R.id.E35);
        e36 = findViewById(R.id.E36);
        e37 = findViewById(R.id.E37);
        e38 = findViewById(R.id.E38);
        e39 = findViewById(R.id.E39);
        e40 = findViewById(R.id.E40);
        e41 = findViewById(R.id.E41);
        e42 = findViewById(R.id.E42);
        e43 = findViewById(R.id.E43);
        e44 = findViewById(R.id.E44);
        e45 = findViewById(R.id.E45);
        e46 = findViewById(R.id.E46);
        e47 = findViewById(R.id.E47);
        e48 = findViewById(R.id.E48);
        e49 = findViewById(R.id.E49);
        e50 = findViewById(R.id.E50);
        e51 = findViewById(R.id.E51);
        e52 = findViewById(R.id.E52);
        e53 = findViewById(R.id.E53);
        e54 = findViewById(R.id.E54);
        e55 = findViewById(R.id.E55);
        e56 = findViewById(R.id.E56);
        e57 = findViewById(R.id.E57);
        e58 = findViewById(R.id.E58);
        e59 = findViewById(R.id.E59);
        e60 = findViewById(R.id.E60);
        e61 = findViewById(R.id.E61);
        e62 = findViewById(R.id.E62);
        e63 = findViewById(R.id.E63);
        e64 = findViewById(R.id.E64);
        e65 = findViewById(R.id.E65);
        e66 = findViewById(R.id.E66);
        e67 = findViewById(R.id.E67);
        e68 = findViewById(R.id.E68);
        e69 = findViewById(R.id.E69);
        e70 = findViewById(R.id.E70);
        e71 = findViewById(R.id.E71);
        e72 = findViewById(R.id.E72);
        e73 = findViewById(R.id.E73);
        e74 = findViewById(R.id.E74);
        e75 = findViewById(R.id.E75);
        e76 = findViewById(R.id.E76);
        e77 = findViewById(R.id.E77);
        e78 = findViewById(R.id.E78);
        e79 = findViewById(R.id.E79);
        e80 = findViewById(R.id.E80);
        e81 = findViewById(R.id.E81);
        e82 = findViewById(R.id.E82);
        e83 = findViewById(R.id.E83);
        e84 = findViewById(R.id.E84);
        e85 = findViewById(R.id.E85);
        e86 = findViewById(R.id.E86);
        e87 = findViewById(R.id.E87);
        e88 = findViewById(R.id.E88);
        e89 = findViewById(R.id.E89);
        e90 = findViewById(R.id.E90);
        e91 = findViewById(R.id.E91);
        e92 = findViewById(R.id.E92);
        e93 = findViewById(R.id.E93);


        header = findViewById(R.id.HeaderC);

        Bundle Uno = getIntent().getExtras();
        UnitNo = Uno.getString("uno");
        header.setText("Needs and Shortage of Unit " + UnitNo);


        save = findViewById(R.id.Save);


        ReadDataFromExcel(Shortage.this, "Unit " + UnitNo + " Shortage.xls");

        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveDatainxcel(Shortage.this, "Unit " + UnitNo + " Shortage.xls");

            }
        });

    }

    private void ReadDataFromExcel(Shortage Shortage, String filename) {

        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook shrt = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet BUS_SYSTEM = shrt.getSheetAt(0);

            HSSFRow Condenser = BUS_SYSTEM.getRow(8);
            HSSFCell ConCell = Condenser.getCell(0);
            e1.setText(ConCell.toString());
            HSSFRow Condense = BUS_SYSTEM.getRow(9);
            HSSFCell ConCel2 = Condense.getCell(0);
            e2.setText(ConCel2.toString());
            HSSFRow Condens = BUS_SYSTEM.getRow(10);
            HSSFCell ConCel3 = Condens.getCell(0);
            e3.setText(ConCel3.toString());
            HSSFRow Conde = BUS_SYSTEM.getRow(11);
            HSSFCell ConCel4 = Conde.getCell(0);
            e4.setText(ConCel4.toString());
            HSSFRow Cond = BUS_SYSTEM.getRow(12);
            HSSFCell ConCel5 = Cond.getCell(0);
            e5.setText(ConCel5.toString());
            HSSFRow Con = BUS_SYSTEM.getRow(13);
            HSSFCell ConCel6 = Con.getCell(0);
            e6.setText(ConCel6.toString());
            HSSFRow Condd = BUS_SYSTEM.getRow(14);
            HSSFCell ConCel7 = Condd.getCell(0);
            e7.setText(ConCel7.toString());
            HSSFRow Condee = BUS_SYSTEM.getRow(15);
            HSSFCell ConCel8 = Condee.getCell(0);
            e8.setText(ConCel8.toString());
            HSSFRow Conn = BUS_SYSTEM.getRow(16);
            HSSFCell ConCel9 = Conn.getCell(0);
            e9.setText(ConCel9.toString());
            HSSFRow Connn = BUS_SYSTEM.getRow(17);
            HSSFCell ConCell0 = Connn.getCell(0);
            e10.setText(ConCell0.toString());
            HSSFRow Condeee = BUS_SYSTEM.getRow(18);
            HSSFCell ConCell1 = Condeee.getCell(0);
            e11.setText(ConCell1.toString());
            HSSFRow Conddd = BUS_SYSTEM.getRow(19);
            HSSFCell ConCell2 = Conddd.getCell(0);
            e12.setText(ConCell2.toString());
            HSSFRow Coo = BUS_SYSTEM.getRow(20);
            HSSFCell ConCell3 = Coo.getCell(0);
            e13.setText(ConCell3.toString());
            HSSFRow Cooo = BUS_SYSTEM.getRow(21);
            HSSFCell ConCell4 = Cooo.getCell(0);
            e14.setText(ConCell4.toString());
            HSSFRow Connnn = BUS_SYSTEM.getRow(22);
            HSSFCell ConCell5 = Connnn.getCell(0);
            e15.setText(ConCell5.toString());
            HSSFRow Cc = BUS_SYSTEM.getRow(23);
            HSSFCell ConCell6 = Cc.getCell(0);
            e16.setText(ConCell6.toString());
            HSSFRow Ccc = BUS_SYSTEM.getRow(24);
            HSSFCell ConCell7 = Ccc.getCell(0);
            e17.setText(ConCell7.toString());
            HSSFRow Cccc = BUS_SYSTEM.getRow(25);
            HSSFCell ConCell8 = Cccc.getCell(0);
            e18.setText(ConCell8.toString());
            HSSFRow Coooo = BUS_SYSTEM.getRow(26);
            HSSFCell ConCell9 = Coooo.getCell(0);
            e19.setText(ConCell9.toString());
            HSSFRow sss = BUS_SYSTEM.getRow(27);
            HSSFCell ConCe20 = sss.getCell(0);
            e20.setText(ConCe20.toString());
            HSSFRow ssss = BUS_SYSTEM.getRow(28);
            HSSFCell ConCe21 = ssss.getCell(0);

            e21.setText(ConCe21.toString());
            HSSFRow xxx = BUS_SYSTEM.getRow(29);
            HSSFCell ConCe22 = xxx.getCell(0);
            e22.setText(ConCe22.toString());
            HSSFRow xxxx = BUS_SYSTEM.getRow(30);
            HSSFCell ConCe23 = xxxx.getCell(0);
            e23.setText(ConCe23.toString());
            HSSFRow xxxxx = BUS_SYSTEM.getRow(31);
            HSSFCell ConCe24 = xxxxx.getCell(0);
            e24.setText(ConCe24.toString());
            HSSFRow zz = BUS_SYSTEM.getRow(32);
            HSSFCell ConCe25 = zz.getCell(0);
            e25.setText(ConCe25.toString());
            HSSFRow zzz = BUS_SYSTEM.getRow(33);
            HSSFCell ConCe26 = zzz.getCell(0);
            e26.setText(ConCe26.toString());
            HSSFRow vv = BUS_SYSTEM.getRow(34);
            HSSFCell ConCe27 = vv.getCell(0);
            e27.setText(ConCe27.toString());
            HSSFRow vvv = BUS_SYSTEM.getRow(35);
            HSSFCell ConCe28 = vvv.getCell(0);
            e28.setText(ConCe28.toString());
            HSSFRow vvvv = BUS_SYSTEM.getRow(36);
            HSSFCell ConCe29 = vvvv.getCell(0);
            e29.setText(ConCe29.toString());
            HSSFRow vvvvv = BUS_SYSTEM.getRow(37);
            HSSFCell ConCe30 = vvvvv.getCell(0);
            e30.setText(ConCe30.toString());
            HSSFRow bb = BUS_SYSTEM.getRow(38);
            HSSFCell ConCe31 = bb.getCell(0);
            e31.setText(ConCe31.toString());


            
            HSSFRow bbb = BUS_SYSTEM.getRow(8);
            HSSFCell ConCe32 = bbb.getCell(3);
            e32.setText(ConCe32.toString());
            HSSFRow bbbb = BUS_SYSTEM.getRow(9);
            HSSFCell ConCe33 = bbbb.getCell(3);
            e33.setText(ConCe33.toString());
            HSSFRow nn = BUS_SYSTEM.getRow(10);
            HSSFCell ConCe34 = nn.getCell(3);
            e34.setText(ConCe34.toString());
            HSSFRow qq = BUS_SYSTEM.getRow(11);
            HSSFCell ConCe35 = qq.getCell(3);
            e35.setText(ConCe35.toString());
            HSSFRow qqq = BUS_SYSTEM.getRow(12);
            HSSFCell ConCe36 = qqq.getCell(3);
            e36.setText(ConCe36.toString());
            HSSFRow qqqq = BUS_SYSTEM.getRow(13);
            HSSFCell ConCe37 = qqqq.getCell(3);
            e37.setText(ConCe37.toString());
            HSSFRow ww = BUS_SYSTEM.getRow(14);
            HSSFCell ConCe38 = ww.getCell(3);
            e38.setText(ConCe38.toString());
            HSSFRow www = BUS_SYSTEM.getRow(15);
            HSSFCell ConCe39 = www.getCell(3);
            e39.setText(ConCe39.toString());
            HSSFRow wwww = BUS_SYSTEM.getRow(16);
            HSSFCell ConCe40 = wwww.getCell(3);
            e40.setText(ConCe40.toString());
            HSSFRow wwwww = BUS_SYSTEM.getRow(17);
            HSSFCell ConCe41 = wwwww.getCell(3);
            e41.setText(ConCe41.toString());
            HSSFRow ee = BUS_SYSTEM.getRow(18);
            HSSFCell ConCe42 = ee.getCell(3);
            e42.setText(ConCe42.toString());
            HSSFRow eee = BUS_SYSTEM.getRow(19);
            HSSFCell ConCe43 = eee.getCell(3);
            e43.setText(ConCe43.toString());
            HSSFRow eeee = BUS_SYSTEM.getRow(20);
            HSSFCell ConCe44 = eeee.getCell(3);
            e44.setText(ConCe44.toString());
            HSSFRow rr = BUS_SYSTEM.getRow(21);
            HSSFCell ConCe45 = rr.getCell(3);
            e45.setText(ConCe45.toString());
            HSSFRow rrr = BUS_SYSTEM.getRow(22);
            HSSFCell ConCe46 = rrr.getCell(3);
            e46.setText(ConCe46.toString());
            HSSFRow rrrr = BUS_SYSTEM.getRow(23);
            HSSFCell ConCe47 = rrrr.getCell(3);
            e47.setText(ConCe47.toString());
            HSSFRow rrrrr = BUS_SYSTEM.getRow(24);
            HSSFCell ConCe48 = rrrrr.getCell(3);
            e48.setText(ConCe48.toString());
            HSSFRow tt = BUS_SYSTEM.getRow(25);
            HSSFCell ConCe49 = tt.getCell(3);
            e49.setText(ConCe49.toString());
            HSSFRow ttt = BUS_SYSTEM.getRow(26);
            HSSFCell ConCe50 = ttt.getCell(3);
            e50.setText(ConCe50.toString());
            HSSFRow tttt = BUS_SYSTEM.getRow(27);
            HSSFCell ConCe550 = tttt.getCell(3);
            e51.setText(ConCe550.toString());
            HSSFRow ttttt = BUS_SYSTEM.getRow(28);
            HSSFCell ConCe504 = ttttt.getCell(3);
            e52.setText(ConCe504.toString());
            HSSFRow tttq = BUS_SYSTEM.getRow(29);
            HSSFCell ConCe505 = tttq.getCell(3);
            e53.setText(ConCe505.toString());
            HSSFRow tttqq = BUS_SYSTEM.getRow(30);
            HSSFCell ConCe506 = tttqq.getCell(3);
            e54.setText(ConCe506.toString());
            HSSFRow tttqqq = BUS_SYSTEM.getRow(31);
            HSSFCell ConCe507 = tttqqq.getCell(3);
            e55.setText(ConCe507.toString());
            HSSFRow tttw = BUS_SYSTEM.getRow(32);
            HSSFCell ConCe508 = tttw.getCell(3);
            e56.setText(ConCe508.toString());
            HSSFRow tttww = BUS_SYSTEM.getRow(33);
            HSSFCell ConCe509 = tttww.getCell(3);
            e57.setText(ConCe509.toString());
            HSSFRow tttwww = BUS_SYSTEM.getRow(34);
            HSSFCell ConCe510 = tttwww.getCell(3);
            e58.setText(ConCe510.toString());
            HSSFRow ttte = BUS_SYSTEM.getRow(35);
            HSSFCell ConCe511 = ttte.getCell(3);
            e59.setText(ConCe511.toString());
            HSSFRow tttee = BUS_SYSTEM.getRow(36);
            HSSFCell ConCe512 = tttee.getCell(3);
            e60.setText(ConCe512.toString());
            HSSFRow ttteee = BUS_SYSTEM.getRow(37);
            HSSFCell ConCe513 = ttteee.getCell(3);
            e61.setText(ConCe513.toString());
            HSSFRow tttr = BUS_SYSTEM.getRow(38);
            HSSFCell ConCe514 = tttr.getCell(3);
            e62.setText(ConCe514.toString());



            HSSFRow bbbq = BUS_SYSTEM.getRow(8);
            HSSFCell ConCe515 = bbbq.getCell(6);
            e63.setText(ConCe515.toString());
            HSSFRow bbbbq = BUS_SYSTEM.getRow(9);
            HSSFCell ConCe516 = bbbbq.getCell(6);
            e64.setText(ConCe516.toString());
            HSSFRow nn1 = BUS_SYSTEM.getRow(10);
            HSSFCell ConCe341 = nn1.getCell(6);
            e65.setText(ConCe341.toString());
            HSSFRow qq1 = BUS_SYSTEM.getRow(11);
            HSSFCell ConCe351 = qq1.getCell(6);
            e66.setText(ConCe351.toString());
            HSSFRow qqq1 = BUS_SYSTEM.getRow(12);
            HSSFCell ConCe361 = qqq1.getCell(6);
            e67.setText(ConCe361.toString());
            HSSFRow qqqq1 = BUS_SYSTEM.getRow(13);
            HSSFCell ConCe371 = qqqq1.getCell(6);
            e68.setText(ConCe371.toString());
            HSSFRow ww1 = BUS_SYSTEM.getRow(14);
            HSSFCell ConCe381 = ww1.getCell(6);
            e69.setText(ConCe381.toString());
            HSSFRow www1 = BUS_SYSTEM.getRow(15);
            HSSFCell ConCe391 = www1.getCell(6);
            e70.setText(ConCe391.toString());
            HSSFRow wwww1 = BUS_SYSTEM.getRow(16);
            HSSFCell ConCe401 = wwww1.getCell(6);
            e71.setText(ConCe401.toString());
            HSSFRow wwwww1 = BUS_SYSTEM.getRow(17);
            HSSFCell ConCe411 = wwwww1.getCell(6);
            e72.setText(ConCe411.toString());
            HSSFRow ee1 = BUS_SYSTEM.getRow(18);
            HSSFCell ConCe421 = ee1.getCell(6);
            e73.setText(ConCe421.toString());
            HSSFRow eee1 = BUS_SYSTEM.getRow(19);
            HSSFCell ConCe431 = eee1.getCell(6);
            e74.setText(ConCe431.toString());
            HSSFRow eeee1 = BUS_SYSTEM.getRow(20);
            HSSFCell ConCe441 = eeee1.getCell(6);
            e75.setText(ConCe441.toString());
            HSSFRow rr1 = BUS_SYSTEM.getRow(21);
            HSSFCell ConCe451 = rr1.getCell(6);
            e76.setText(ConCe451.toString());
            HSSFRow rrr1 = BUS_SYSTEM.getRow(22);
            HSSFCell ConCe461 = rrr1.getCell(6);
            e77.setText(ConCe461.toString());
            HSSFRow rrrr1 = BUS_SYSTEM.getRow(23);
            HSSFCell ConCe471 = rrrr1.getCell(6);
            e78.setText(ConCe471.toString());
            HSSFRow rrrrr1 = BUS_SYSTEM.getRow(24);
            HSSFCell ConCe481 = rrrrr1.getCell(6);
            e79.setText(ConCe481.toString());
            HSSFRow tt1 = BUS_SYSTEM.getRow(25);
            HSSFCell ConCe491 = tt1.getCell(6);
            e80.setText(ConCe491.toString());
            HSSFRow ttt1 = BUS_SYSTEM.getRow(26);
            HSSFCell ConCe501 = ttt1.getCell(6);
            e81.setText(ConCe501.toString());
            HSSFRow tttt1 = BUS_SYSTEM.getRow(27);
            HSSFCell ConCe5501 = tttt1.getCell(6);
            e82.setText(ConCe5501.toString());
            HSSFRow ttttt1 = BUS_SYSTEM.getRow(28);
            HSSFCell ConCe5041 = ttttt1.getCell(6);
            e83.setText(ConCe5041.toString());
            HSSFRow tttq1 = BUS_SYSTEM.getRow(29);
            HSSFCell ConCe5051 = tttq1.getCell(6);
            e84.setText(ConCe5051.toString());
            HSSFRow tttqq1 = BUS_SYSTEM.getRow(30);
            HSSFCell ConCe5061 = tttqq1.getCell(6);
            e85.setText(ConCe5061.toString());
            HSSFRow tttqqq1 = BUS_SYSTEM.getRow(31);
            HSSFCell ConCe5071 = tttqqq1.getCell(6);
            e86.setText(ConCe5071.toString());
            HSSFRow tttw1 = BUS_SYSTEM.getRow(32);
            HSSFCell ConCe5081 = tttw1.getCell(6);
            e87.setText(ConCe5081.toString());
            HSSFRow tttww1 = BUS_SYSTEM.getRow(33);
            HSSFCell ConCe5091 = tttww1.getCell(6);
            e88.setText(ConCe5091.toString());
            HSSFRow tttwww1 = BUS_SYSTEM.getRow(34);
            HSSFCell ConCe5101 = tttwww1.getCell(6);
            e89.setText(ConCe5101.toString());
            HSSFRow ttte1 = BUS_SYSTEM.getRow(35);
            HSSFCell ConCe5111 = ttte1.getCell(6);
            e90.setText(ConCe5111.toString());
            HSSFRow tttee1 = BUS_SYSTEM.getRow(36);
            HSSFCell ConCe5121 = tttee1.getCell(6);
            e91.setText(ConCe5121.toString());
            HSSFRow ttteee1 = BUS_SYSTEM.getRow(37);
            HSSFCell ConCe5131 = ttteee1.getCell(6);
            e92.setText(ConCe5131.toString());
            HSSFRow tttr1 = BUS_SYSTEM.getRow(38);
            HSSFCell ConCe5141 = tttr1.getCell(6);
            e93.setText(ConCe5141.toString());





        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void SaveDatainxcel(Context context, String filename) {

        try {
            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);

            HSSFWorkbook workbook = new HSSFWorkbook(myInputStream);
            HSSFSheet BUS_SYSTEM = workbook.getSheetAt(0);

            HSSFRow a1 = BUS_SYSTEM.getRow(8);
            HSSFCell con1 = a1.getCell(0);
            con1.setCellValue(e1.getText().toString().trim());

            HSSFRow a2 = BUS_SYSTEM.getRow(9);
            HSSFCell con2 = a2.getCell(0);
            con2.setCellValue(e2.getText().toString().trim());

            HSSFRow a3 = BUS_SYSTEM.getRow(10);
            HSSFCell con3 = a3.getCell(0);
            con3.setCellValue(e3.getText().toString().trim());

            HSSFRow a4 = BUS_SYSTEM.getRow(11);
            HSSFCell con4 = a4.getCell(0);
            con4.setCellValue(e4.getText().toString().trim());

            HSSFRow a5 = BUS_SYSTEM.getRow(12);
            HSSFCell con5 = a5.getCell(0);
            con5.setCellValue(e5.getText().toString().trim());

            HSSFRow a6 = BUS_SYSTEM.getRow(13);
            HSSFCell con6 = a6.getCell(0);
            con6.setCellValue(e6.getText().toString().trim());

            HSSFRow a7 = BUS_SYSTEM.getRow(14);
            HSSFCell con7 = a7.getCell(0);
            con7.setCellValue(e7.getText().toString().trim());

            HSSFRow a8 = BUS_SYSTEM.getRow(15);
            HSSFCell con8 = a8.getCell(0);
            con8.setCellValue(e8.getText().toString().trim());

            HSSFRow a9 = BUS_SYSTEM.getRow(16);
            HSSFCell con9 = a9.getCell(0);
            con9.setCellValue(e9.getText().toString().trim());

            HSSFRow a10 = BUS_SYSTEM.getRow(17);
            HSSFCell con10 = a10.getCell(0);
            con10.setCellValue(e10.getText().toString().trim());

            HSSFRow a11 = BUS_SYSTEM.getRow(18);
            HSSFCell con11 = a11.getCell(0);
            con11.setCellValue(e11.getText().toString().trim());

            HSSFRow a12 = BUS_SYSTEM.getRow(19);
            HSSFCell con12 = a12.getCell(0);
            con12.setCellValue(e12.getText().toString().trim());

            HSSFRow a13 = BUS_SYSTEM.getRow(20);
            HSSFCell con13 = a13.getCell(0);
            con13.setCellValue(e13.getText().toString().trim());

            HSSFRow a14 = BUS_SYSTEM.getRow(21);
            HSSFCell con14 = a14.getCell(0);
            con14.setCellValue(e14.getText().toString().trim());

            HSSFRow a15 = BUS_SYSTEM.getRow(22);
            HSSFCell con15 = a15.getCell(0);
            con15.setCellValue(e15.getText().toString().trim());

            HSSFRow a16 = BUS_SYSTEM.getRow(23);
            HSSFCell con16 = a16.getCell(0);
            con16.setCellValue(e16.getText().toString().trim());

            HSSFRow a17 = BUS_SYSTEM.getRow(24);
            HSSFCell con17 = a17.getCell(0);
            con17.setCellValue(e17.getText().toString().trim());

            HSSFRow a18 = BUS_SYSTEM.getRow(25);
            HSSFCell con18 = a18.getCell(0);
            con18.setCellValue(e18.getText().toString().trim());

            HSSFRow a19 = BUS_SYSTEM.getRow(26);
            HSSFCell con19 = a19.getCell(0);
            con19.setCellValue(e19.getText().toString().trim());

            HSSFRow a20 = BUS_SYSTEM.getRow(27);
            HSSFCell con20 = a20.getCell(0);
            con20.setCellValue(e20.getText().toString().trim());

            HSSFRow a21 = BUS_SYSTEM.getRow(28);
            HSSFCell con21 = a21.getCell(0);
            con21.setCellValue(e21.getText().toString().trim());

            HSSFRow a22 = BUS_SYSTEM.getRow(29);
            HSSFCell con22 = a22.getCell(0);
            con22.setCellValue(e22.getText().toString().trim());

            HSSFRow a23 = BUS_SYSTEM.getRow(30);
            HSSFCell con23 = a23.getCell(0);
            con23.setCellValue(e23.getText().toString().trim());

            HSSFRow a24 = BUS_SYSTEM.getRow(31);
            HSSFCell con24 = a24.getCell(0);
            con24.setCellValue(e24.getText().toString().trim());

            HSSFRow a25 = BUS_SYSTEM.getRow(32);
            HSSFCell con25 = a25.getCell(0);
            con25.setCellValue(e25.getText().toString().trim());

            HSSFRow a26 = BUS_SYSTEM.getRow(33);
            HSSFCell con26 = a26.getCell(0);
            con26.setCellValue(e26.getText().toString().trim());

            HSSFRow a27 = BUS_SYSTEM.getRow(34);
            HSSFCell con27 = a27.getCell(0);
            con27.setCellValue(e27.getText().toString().trim());

            HSSFRow a28 = BUS_SYSTEM.getRow(35);
            HSSFCell con28 = a28.getCell(0);
            con28.setCellValue(e28.getText().toString().trim());

            HSSFRow a29 = BUS_SYSTEM.getRow(36);
            HSSFCell con29 = a29.getCell(0);
            con29.setCellValue(e29.getText().toString().trim());

            HSSFRow a30 = BUS_SYSTEM.getRow(37);
            HSSFCell con30 = a30.getCell(0);
            con30.setCellValue(e30.getText().toString().trim());

            HSSFRow a31 = BUS_SYSTEM.getRow(38);
            HSSFCell con31 = a31.getCell(0);
            con31.setCellValue(e31.getText().toString().trim());

            HSSFRow a32 = BUS_SYSTEM.getRow(8);
            HSSFCell con32 = a32.getCell(3);
            con32.setCellValue(e32.getText().toString().trim());

            HSSFRow a33 = BUS_SYSTEM.getRow(9);
            HSSFCell con33 = a33.getCell(3);
            con33.setCellValue(e33.getText().toString().trim());

            HSSFRow a34 = BUS_SYSTEM.getRow(10);
            HSSFCell con34 = a34.getCell(3);
            con34.setCellValue(e34.getText().toString().trim());

            HSSFRow a35 = BUS_SYSTEM.getRow(11);
            HSSFCell con35 = a35.getCell(3);
            con35.setCellValue(e35.getText().toString().trim());

            HSSFRow a36 = BUS_SYSTEM.getRow(12);
            HSSFCell con36 = a36.getCell(3);
            con36.setCellValue(e36.getText().toString().trim());

            HSSFRow a37 = BUS_SYSTEM.getRow(13);
            HSSFCell con37 = a37.getCell(3);
            con37.setCellValue(e37.getText().toString().trim());

            HSSFRow a38 = BUS_SYSTEM.getRow(14);
            HSSFCell con38 = a38.getCell(3);
            con38.setCellValue(e38.getText().toString().trim());

            HSSFRow a39 = BUS_SYSTEM.getRow(15);
            HSSFCell con39 = a39.getCell(3);
            con39.setCellValue(e39.getText().toString().trim());

            HSSFRow a40 = BUS_SYSTEM.getRow(16);
            HSSFCell con40 = a40.getCell(3);
            con40.setCellValue(e40.getText().toString().trim());

            HSSFRow a41 = BUS_SYSTEM.getRow(17);
            HSSFCell con41 = a41.getCell(3);
            con41.setCellValue(e41.getText().toString().trim());

            HSSFRow a42 = BUS_SYSTEM.getRow(18);
            HSSFCell con42 = a42.getCell(3);
            con42.setCellValue(e42.getText().toString().trim());

            HSSFRow a43 = BUS_SYSTEM.getRow(19);
            HSSFCell con43 = a43.getCell(3);
            con43.setCellValue(e43.getText().toString().trim());

            HSSFRow a44 = BUS_SYSTEM.getRow(20);
            HSSFCell con44 = a44.getCell(3);
            con44.setCellValue(e44.getText().toString().trim());

            HSSFRow a45 = BUS_SYSTEM.getRow(21);
            HSSFCell con45 = a45.getCell(3);
            con45.setCellValue(e45.getText().toString().trim());

            HSSFRow a46 = BUS_SYSTEM.getRow(22);
            HSSFCell con46 = a46.getCell(3);
            con46.setCellValue(e46.getText().toString().trim());

            HSSFRow a47 = BUS_SYSTEM.getRow(23);
            HSSFCell con47 = a47.getCell(3);
            con47.setCellValue(e47.getText().toString().trim());

            HSSFRow a48 = BUS_SYSTEM.getRow(24);
            HSSFCell con48 = a48.getCell(3);
            con48.setCellValue(e48.getText().toString().trim());

            HSSFRow a49 = BUS_SYSTEM.getRow(25);
            HSSFCell con49 = a49.getCell(3);
            con49.setCellValue(e49.getText().toString().trim());

            HSSFRow a50 = BUS_SYSTEM.getRow(26);
            HSSFCell con50 = a50.getCell(3);
            con50.setCellValue(e50.getText().toString().trim());

            HSSFRow a51 = BUS_SYSTEM.getRow(27);
            HSSFCell con461 = a51.getCell(3);
            con461.setCellValue(e51.getText().toString().trim());

            HSSFRow a471 = BUS_SYSTEM.getRow(28);
            HSSFCell con471 = a471.getCell(3);
            con471.setCellValue(e52.getText().toString().trim());

            HSSFRow a481 = BUS_SYSTEM.getRow(29);
            HSSFCell con481 = a481.getCell(3);
            con481.setCellValue(e53.getText().toString().trim());

            HSSFRow a491 = BUS_SYSTEM.getRow(30);
            HSSFCell con491 = a491.getCell(3);
            con491.setCellValue(e54.getText().toString().trim());

            HSSFRow a501 = BUS_SYSTEM.getRow(31);
            HSSFCell con501 = a501.getCell(3);
            con501.setCellValue(e55.getText().toString().trim());


            HSSFRow a492 = BUS_SYSTEM.getRow(32);
            HSSFCell con492 = a492.getCell(3);
            con492.setCellValue(e56.getText().toString().trim());

            HSSFRow a502 = BUS_SYSTEM.getRow(33);
            HSSFCell con502 = a502.getCell(3);
            con502.setCellValue(e57.getText().toString().trim());

            HSSFRow a512 = BUS_SYSTEM.getRow(34);
            HSSFCell con4612 = a512.getCell(3);
            con4612.setCellValue(e58.getText().toString().trim());

            HSSFRow a4712 = BUS_SYSTEM.getRow(35);
            HSSFCell con4712 = a4712.getCell(3);
            con4712.setCellValue(e59.getText().toString().trim());

            HSSFRow a4812 = BUS_SYSTEM.getRow(36);
            HSSFCell con4812 = a4812.getCell(3);
            con4812.setCellValue(e60.getText().toString().trim());

            HSSFRow a4912 = BUS_SYSTEM.getRow(37);
            HSSFCell con4912 = a4912.getCell(3);
            con4912.setCellValue(e61.getText().toString().trim());

            HSSFRow a5012 = BUS_SYSTEM.getRow(38);
            HSSFCell con5012 = a5012.getCell(3);
            con5012.setCellValue(e62.getText().toString().trim());

            HSSFRow a50123 = BUS_SYSTEM.getRow(8);
            HSSFCell con50123 = a50123.getCell(6);
            con50123.setCellValue(e63.getText().toString().trim());

            HSSFRow a501231 = BUS_SYSTEM.getRow(9);
            HSSFCell con501231 = a501231.getCell(6);
            con501231.setCellValue(e64.getText().toString().trim());

            HSSFRow a5012311 = BUS_SYSTEM.getRow(10);
            HSSFCell con5012311 = a5012311.getCell(6);
            con5012311.setCellValue(e65.getText().toString().trim());
            HSSFRow a501232 = BUS_SYSTEM.getRow(11);
            HSSFCell con501232 = a501232.getCell(6);
            con501232.setCellValue(e66.getText().toString().trim());
            HSSFRow a501234 = BUS_SYSTEM.getRow(12);
            HSSFCell con501234 = a501234.getCell(6);
            con501234.setCellValue(e67.getText().toString().trim());
            HSSFRow a501235 = BUS_SYSTEM.getRow(13);
            HSSFCell con501235 = a501235.getCell(6);
            con501235.setCellValue(e68.getText().toString().trim());
            HSSFRow a501236 = BUS_SYSTEM.getRow(14);
            HSSFCell con501236 = a501236.getCell(6);
            con501236.setCellValue(e69.getText().toString().trim());
            HSSFRow a501237 = BUS_SYSTEM.getRow(15);
            HSSFCell con501237 = a501237.getCell(6);
            con501237.setCellValue(e70.getText().toString().trim());
            HSSFRow a501238 = BUS_SYSTEM.getRow(16);
            HSSFCell con501238 = a501238.getCell(6);
            con501238.setCellValue(e71.getText().toString().trim());
            HSSFRow a501239 = BUS_SYSTEM.getRow(17);
            HSSFCell con501239 = a501239.getCell(6);
            con501239.setCellValue(e72.getText().toString().trim());
            HSSFRow a501230 = BUS_SYSTEM.getRow(18);
            HSSFCell con501230 = a501230.getCell(6);
            con501230.setCellValue(e73.getText().toString().trim());
            HSSFRow a5012322 = BUS_SYSTEM.getRow(19);
            HSSFCell con5012322 = a5012322.getCell(6);
            con5012322.setCellValue(e74.getText().toString().trim());
            HSSFRow a5012333 = BUS_SYSTEM.getRow(20);
            HSSFCell con5012333 = a5012333.getCell(6);
            con5012333.setCellValue(e75.getText().toString().trim());
            HSSFRow a50123z = BUS_SYSTEM.getRow(21);
            HSSFCell con50123z = a50123z.getCell(6);
            con50123z.setCellValue(e76.getText().toString().trim());
            HSSFRow a50123x = BUS_SYSTEM.getRow(22);
            HSSFCell con50123x = a50123x.getCell(6);
            con50123x.setCellValue(e77.getText().toString().trim());
            HSSFRow a50123c = BUS_SYSTEM.getRow(23);
            HSSFCell con50123c = a50123c.getCell(6);
            con50123c.setCellValue(e78.getText().toString().trim());
            HSSFRow a50123v = BUS_SYSTEM.getRow(24);
            HSSFCell con50123v = a50123v.getCell(6);
            con50123v.setCellValue(e79.getText().toString().trim());
            HSSFRow a50123vv = BUS_SYSTEM.getRow(25);
            HSSFCell con50123vv = a50123vv.getCell(6);
            con50123vv.setCellValue(e80.getText().toString().trim());
            HSSFRow a50123o = BUS_SYSTEM.getRow(26);
            HSSFCell con50123o = a50123o.getCell(6);
            con50123o.setCellValue(e81.getText().toString().trim());
            HSSFRow a50123oo = BUS_SYSTEM.getRow(27);
            HSSFCell con50123oo = a50123oo.getCell(6);
            con50123oo.setCellValue(e82.getText().toString().trim());
            HSSFRow a50123oo1 = BUS_SYSTEM.getRow(28);
            HSSFCell con50123oo1 = a50123oo1.getCell(6);
            con50123oo1.setCellValue(e83.getText().toString().trim());
            HSSFRow a50123oow = BUS_SYSTEM.getRow(29);
            HSSFCell con50123oow = a50123oow.getCell(6);
            con50123oow.setCellValue(e84.getText().toString().trim());
            HSSFRow a50123ooe = BUS_SYSTEM.getRow(30);
            HSSFCell con50123ooe = a50123ooe.getCell(6);
            con50123ooe.setCellValue(e85.getText().toString().trim());

            HSSFRow a50123oor = BUS_SYSTEM.getRow(31);
            HSSFCell con50123oor = a50123oor.getCell(6);
            con50123oor.setCellValue(e86.getText().toString().trim());

            HSSFRow a50123oot = BUS_SYSTEM.getRow(32);
            HSSFCell con50123oot = a50123oot.getCell(6);
            con50123oot.setCellValue(e87.getText().toString().trim());

            HSSFRow a50123ooz = BUS_SYSTEM.getRow(33);
            HSSFCell con50123ooz = a50123ooz.getCell(6);
            con50123ooz.setCellValue(e88.getText().toString().trim());

            HSSFRow a50123oox = BUS_SYSTEM.getRow(34);
            HSSFCell con50123oox = a50123oox.getCell(6);
            con50123oox.setCellValue(e89.getText().toString().trim());

            HSSFRow a50123ooc = BUS_SYSTEM.getRow(35);
            HSSFCell con50123ooc = a50123ooc.getCell(6);
            con50123ooc.setCellValue(e90.getText().toString().trim());

            HSSFRow a50123oov = BUS_SYSTEM.getRow(36);
            HSSFCell con50123oov = a50123oov.getCell(6);
            con50123oov.setCellValue(e91.getText().toString().trim());

            HSSFRow a50123oob = BUS_SYSTEM.getRow(37);
            HSSFCell con50123oob = a50123oob.getCell(6);
            con50123oob.setCellValue(e92.getText().toString().trim());

            HSSFRow a50123oon = BUS_SYSTEM.getRow(38);
            HSSFCell con50123oon = a50123oon.getCell(6);
            con50123oon.setCellValue(e93.getText().toString().trim());


            myInputStream.close();
            FileOutputStream fos = new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename));
            workbook.write(fos);
            fos.close();
            Toast.makeText(context, "Changes are done.", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }
}




