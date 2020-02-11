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

public class Stationary extends AppCompatActivity {
    EditText

            c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16, c17, c18, c19, c20, c21, c22, c23, c24, c25, c26, c27, c28, c29, c30, c31, c32, c33, c34, c35, c36, c37, c38, c39, c40, c41, c42, c43, c44, c45, c46, c47, c48, c49, c50, c51, c52, c53, c54, c55, c56, c57, c58, c59, c60, c61, c62, c63, c64, c65, c66, c67, c68, c69, c70, c71, c72, c73, c74, c75, c76, c77, c78, c79, c80, c81, c82, c83, c84, c85, c86, c87, c88, c89, c90, c91, c92, c93, c94, c95, c96, c97, c98, c99, c100, c101, c102, c103, c104, c105, c106, c107, c108, c109, c110, c111, c112, c113, c114, c115, c116, c117, c118, c119, c120, c121, c122, c123, c124, c125, c126, c127, c128, c129, c130, c131, c132, c133, c134, c135;
    Button save;
    String UnitNo;
    TextView header;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(Stationary.this);
        setContentView(R.layout.activity_stationary);

        c1 = findViewById(R.id.Co1);
        c2 = findViewById(R.id.Co2);
        c3 = findViewById(R.id.Co3);
        c4 = findViewById(R.id.Co4);
        c5 = findViewById(R.id.Co5);
        c6 = findViewById(R.id.Co6);
        c7 = findViewById(R.id.Co7);
        c8 = findViewById(R.id.Co8);
        c9 = findViewById(R.id.Co9);
        c10 = findViewById(R.id.Co10);
        c11 = findViewById(R.id.Co11);
        c12 = findViewById(R.id.Co12);
        c13 = findViewById(R.id.Co13);
        c14 = findViewById(R.id.Co14);
        c15 = findViewById(R.id.Co15);
        c16 = findViewById(R.id.Co16);
        c17 = findViewById(R.id.Co17);
        c18 = findViewById(R.id.Co18);
        c19 = findViewById(R.id.Co19);
        c20 = findViewById(R.id.Co20);
        c21 = findViewById(R.id.Co21);
        c22 = findViewById(R.id.Co22);
        c23 = findViewById(R.id.Co23);
        c24 = findViewById(R.id.Co24);
        c25 = findViewById(R.id.Co25);
        c26 = findViewById(R.id.Co26);
        c27 = findViewById(R.id.Co27);
        c28 = findViewById(R.id.Co28);
        c29 = findViewById(R.id.Co29);
        c30 = findViewById(R.id.Co30);
        c31 = findViewById(R.id.Co31);
        c32 = findViewById(R.id.Co32);
        c33 = findViewById(R.id.Co33);
        c34 = findViewById(R.id.Co34);
        c35 = findViewById(R.id.Co35);
        c36 = findViewById(R.id.Co36);
        c37 = findViewById(R.id.Co37);
        c38 = findViewById(R.id.Co38);
        c39 = findViewById(R.id.Co39);
        c40 = findViewById(R.id.Co40);
        c41 = findViewById(R.id.Co41);
        c42 = findViewById(R.id.Co42);
        c43 = findViewById(R.id.Co43);
        c44 = findViewById(R.id.Co44);
        c45 = findViewById(R.id.Co45);
        c46 = findViewById(R.id.Co46);
        c47 = findViewById(R.id.Co47);
        c48 = findViewById(R.id.Co48);
        c49 = findViewById(R.id.Co49);
        c50 = findViewById(R.id.Co50);
        c51 = findViewById(R.id.Co51);
        c52 = findViewById(R.id.Co52);
        c53 = findViewById(R.id.Co53);
        c54 = findViewById(R.id.Co54);
        c55 = findViewById(R.id.Co55);
        c56 = findViewById(R.id.Co56);
        c57 = findViewById(R.id.Co57);
        c58 = findViewById(R.id.Co58);
        c59 = findViewById(R.id.Co59);
        c60 = findViewById(R.id.Co60);
        c61 = findViewById(R.id.Co61);
        c62 = findViewById(R.id.Co62);
        c63 = findViewById(R.id.Co63);
        c64 = findViewById(R.id.Co64);
        c65 = findViewById(R.id.Co65);
        c66 = findViewById(R.id.Co66);
        c67 = findViewById(R.id.Co67);
        c68 = findViewById(R.id.Co68);
        c69 = findViewById(R.id.Co69);
        c70 = findViewById(R.id.Co70);
        c71 = findViewById(R.id.Co71);
        c72 = findViewById(R.id.Co72);
        c73 = findViewById(R.id.Co73);
        c74 = findViewById(R.id.Co74);
        c75 = findViewById(R.id.Co75);
        c76 = findViewById(R.id.Co76);
        c77 = findViewById(R.id.Co77);
        c78 = findViewById(R.id.Co78);
        c79 = findViewById(R.id.Co79);
        c80 = findViewById(R.id.Co80);
        c81 = findViewById(R.id.Co81);
        c82 = findViewById(R.id.Co82);
        c83 = findViewById(R.id.Co83);
        c84 = findViewById(R.id.Co84);
        c85 = findViewById(R.id.Co85);
        c86 = findViewById(R.id.Co86);
        c87 = findViewById(R.id.Co87);
        c88 = findViewById(R.id.Co88);
        c89 = findViewById(R.id.Co89);
        c90 = findViewById(R.id.Co90);
        c91 = findViewById(R.id.Co91);
        c92 = findViewById(R.id.Co92);
        c93 = findViewById(R.id.Co93);
        c94 = findViewById(R.id.Co94);
        c95 = findViewById(R.id.Co95);
        c96 = findViewById(R.id.Co96);
        c97 = findViewById(R.id.Co97);
        c98 = findViewById(R.id.Co98);
        c99 = findViewById(R.id.Co99);
        c100 = findViewById(R.id.Co100);
        c101 = findViewById(R.id.Co101);
        c102 = findViewById(R.id.Co102);
        c103 = findViewById(R.id.Co103);
        c104 = findViewById(R.id.Co104);
        c105 = findViewById(R.id.Co105);
        c106 = findViewById(R.id.Co106);
        c107 = findViewById(R.id.Co107);
        c108 = findViewById(R.id.Co108);
        c109 = findViewById(R.id.Co109);
        c110 = findViewById(R.id.Co110);
        c111 = findViewById(R.id.Co111);
        c112 = findViewById(R.id.Co112);
        c113 = findViewById(R.id.Co113);
        c114 = findViewById(R.id.Co114);
        c115 = findViewById(R.id.Co115);
        c116 = findViewById(R.id.Co116);
        c117 = findViewById(R.id.Co117);
        c118 = findViewById(R.id.Co118);
        c119 = findViewById(R.id.Co119);
        c120 = findViewById(R.id.Co120);
        c121 = findViewById(R.id.Co121);
        c122 = findViewById(R.id.Co122);
        c123 = findViewById(R.id.Co123);
        c124 = findViewById(R.id.Co124);
        c125 = findViewById(R.id.Co125);
        c126 = findViewById(R.id.Co126);
        c127 = findViewById(R.id.Co127);
        c128 = findViewById(R.id.Co128);
        c129 = findViewById(R.id.Co129);
        c130 = findViewById(R.id.Co130);
        c131 = findViewById(R.id.Co131);
        c132 = findViewById(R.id.Co132);
        c133 = findViewById(R.id.Co133);
        c134 = findViewById(R.id.Co134);
        c135 = findViewById(R.id.Co135);

        header=findViewById(R.id.HeaderC) ;

        Bundle Uno=getIntent().getExtras();
        UnitNo=Uno.getString("uno");
        header.setText("Stationary of unit "+UnitNo);


        save=findViewById(R.id.Save);


        ReadDataFromExcel(Stationary.this, "Inventory List unit "+UnitNo+".xls");

        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveDatainxcel(Stationary.this, "Inventory List unit "+UnitNo+".xls");

            }
        });


    }
    private void ReadDataFromExcel(Stationary stationary, String filename) {

        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook inv = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet COMMUNICATION = inv.getSheetAt(11);

            HSSFRow Condenser = COMMUNICATION.getRow(5);
            HSSFCell ConCell = Condenser.getCell(4);
            HSSFCell ConCel2 = Condenser.getCell(5);
            HSSFCell ConCel3 = Condenser.getCell(3);
            c1.setText(ConCell.toString());
            c2.setText(ConCel2.toString());
            c3.setText(ConCel3.toString());

            HSSFRow Eva = COMMUNICATION.getRow(6);
            HSSFCell eval = Eva.getCell(4);
            HSSFCell eva2 = Eva.getCell(5);
            HSSFCell eva3 = Eva.getCell(3);
            c4.setText(eval.toString());
            c5.setText(eva2.toString());
            c6.setText(eva3.toString());

            HSSFRow comp = COMMUNICATION.getRow(7);
            HSSFCell com1 = comp.getCell(4);
            HSSFCell com2 = comp.getCell(5);
            HSSFCell com3 = comp.getCell(3);
            c7.setText(com1.toString());
            c8.setText(com2.toString());
            c9.setText(com3.toString());

            HSSFRow mult = COMMUNICATION.getRow(8);
            HSSFCell mul1 = mult.getCell(4);
            HSSFCell mul2 = mult.getCell(5);
            HSSFCell mul3 = mult.getCell(3);
            c10.setText(mul1.toString());
            c11.setText(mul2.toString());
            c12.setText(mul3.toString());

            HSSFRow ligh = COMMUNICATION.getRow(9);
            HSSFCell lig1 = ligh.getCell(4);
            HSSFCell lig2 = ligh.getCell(5);
            HSSFCell lig3 = ligh.getCell(3);
            c13.setText(lig1.toString());
            c14.setText(lig2.toString());
            c15.setText(lig3.toString());

            HSSFRow term = COMMUNICATION.getRow(10);
            HSSFCell ter1 = term.getCell(4);
            HSSFCell ter2 = term.getCell(5);
            HSSFCell ter3 = term.getCell(3);
            c16.setText(ter1.toString());
            c17.setText(ter2.toString());
            c18.setText(ter3.toString());

            HSSFRow fusi = COMMUNICATION.getRow(11);
            HSSFCell fus1 = fusi.getCell(4);
            HSSFCell fus2 = fusi.getCell(5);
            HSSFCell fus3 = fusi.getCell(3);
            c19.setText(fus1.toString());
            c20.setText(fus2.toString());
            c21.setText(fus3.toString());

            HSSFRow powr = COMMUNICATION.getRow(12);
            HSSFCell pow1 = powr.getCell(4);
            HSSFCell pow2 = powr.getCell(5);
            HSSFCell pow3 = powr.getCell(3);
            c22.setText(pow1.toString());
            c23.setText(pow2.toString());
            c24.setText(pow3.toString());

            HSSFRow trns = COMMUNICATION.getRow(13);
            HSSFCell trn1 = trns.getCell(4);
            HSSFCell trn2 = trns.getCell(5);
            HSSFCell trn3 = trns.getCell(3);
            c25.setText(trn1.toString());
            c26.setText(trn2.toString());
            c27.setText(trn3.toString());

            HSSFRow fani = COMMUNICATION.getRow(14);
            HSSFCell fan1 = fani.getCell(4);
            HSSFCell fan2 = fani.getCell(5);
            HSSFCell fan3 = fani.getCell(3);
            c28.setText(fan1.toString());
            c29.setText(fan2.toString());
            c30.setText(fan3.toString());

            HSSFRow upsi = COMMUNICATION.getRow(15);
            HSSFCell ups1 = upsi.getCell(4);
            HSSFCell ups2 = upsi.getCell(5);
            HSSFCell ups3 = upsi.getCell(3);
            c31.setText(ups1.toString());
            c32.setText(ups2.toString());
            c33.setText(ups3.toString());

            HSSFRow purg = COMMUNICATION.getRow(16);
            HSSFCell pur1 = purg.getCell(4);
            HSSFCell pur2 = purg.getCell(5);
            HSSFCell pur3 = purg.getCell(3);
            c34.setText(pur1.toString());
            c35.setText(pur2.toString());
            c36.setText(pur3.toString());

            HSSFRow smok = COMMUNICATION.getRow(17);
            HSSFCell smo1 = smok.getCell(4);
            HSSFCell smo2 = smok.getCell(5);
            HSSFCell smo3 = smok.getCell(3);
            c37.setText(smo1.toString());
            c38.setText(smo2.toString());
            c39.setText(smo3.toString());

            HSSFRow fire = COMMUNICATION.getRow(18);
            HSSFCell fir1 = fire.getCell(4);
            HSSFCell fir2 = fire.getCell(5);
            HSSFCell fir3 = fire.getCell(3);
            c40.setText(fir1.toString());
            c41.setText(fir2.toString());
            c42.setText(fir3.toString());

            HSSFRow toxi = COMMUNICATION.getRow(19);
            HSSFCell tox1 = toxi.getCell(4);
            HSSFCell tox2 = toxi.getCell(5);
            HSSFCell tox3 = toxi.getCell(3);
            c43.setText(tox1.toString());
            c44.setText(tox2.toString());
            c45.setText(tox3.toString());

            HSSFRow flam = COMMUNICATION.getRow(20);
            HSSFCell flm1 = flam.getCell(4);
            HSSFCell flm2 = flam.getCell(5);
            HSSFCell flm3 = flam.getCell(3);
            c46.setText(flm1.toString());
            c47.setText(flm2.toString());
            c48.setText(flm3.toString());

            HSSFRow firs = COMMUNICATION.getRow(21);
            HSSFCell fira1 = firs.getCell(4);
            HSSFCell fira2 = firs.getCell(5);
            HSSFCell fira3 = firs.getCell(3);
            c49.setText(fira1.toString());
            c50.setText(fira2.toString());
            c51.setText(fira3.toString());

            HSSFRow radi = COMMUNICATION.getRow(22);
            HSSFCell rad1 = radi.getCell(4);
            HSSFCell rad2 = radi.getCell(5);
            HSSFCell rad3 = radi.getCell(3);
            c52.setText(rad1.toString());
            c53.setText(rad2.toString());
            c54.setText(rad3.toString());

            HSSFRow refg = COMMUNICATION.getRow(23);
            HSSFCell ref1 = refg.getCell(4);
            HSSFCell ref2 = refg.getCell(5);
            HSSFCell ref3 = refg.getCell(3);
            c55.setText(ref1.toString());
            c56.setText(ref2.toString());
            c57.setText(ref3.toString());

            HSSFRow sink = COMMUNICATION.getRow(24);
            HSSFCell sin1 = sink.getCell(4);
            HSSFCell sin2 = sink.getCell(5);
            HSSFCell sin3 = sink.getCell(3);
            c58.setText(sin1.toString());
            c59.setText(sin2.toString());
            c60.setText(sin3.toString());

            HSSFRow exti = COMMUNICATION.getRow(25);
            HSSFCell ext1 = exti.getCell(4);
            HSSFCell ext2 = exti.getCell(5);
            HSSFCell ext3 = exti.getCell(3);
            c61.setText(ext1.toString());
            c62.setText(ext2.toString());
            c63.setText(ext3.toString());

            HSSFRow chai = COMMUNICATION.getRow(26);
            HSSFCell cha1 = chai.getCell(4);
            HSSFCell cha2 = chai.getCell(5);
            HSSFCell cha3 = chai.getCell(3);
            c64.setText(cha1.toString());
            c65.setText(cha2.toString());
            c66.setText(cha3.toString());

            HSSFRow thrm = COMMUNICATION.getRow(27);
            HSSFCell thr1 = thrm.getCell(4);
            HSSFCell thr2 = thrm.getCell(5);
            HSSFCell thr3 = thrm.getCell(3);
            c67.setText(thr1.toString());
            c68.setText(thr2.toString());
            c69.setText(thr3.toString());

            HSSFRow hydr = COMMUNICATION.getRow(28);
            HSSFCell hyd1 = hydr.getCell(4);
            HSSFCell hyd2 = hydr.getCell(5);
            HSSFCell hyd3 = hydr.getCell(3);
            c70.setText(hyd1.toString());
            c71.setText(hyd2.toString());
            c72.setText(hyd3.toString());

            HSSFRow antn = COMMUNICATION.getRow(29);
            HSSFCell ant1 = antn.getCell(4);
            HSSFCell ant2 = antn.getCell(5);
            HSSFCell ant3 = antn.getCell(3);
            c73.setText(ant1.toString());
            c74.setText(ant2.toString());
            c75.setText(ant3.toString());

            HSSFRow rack = COMMUNICATION.getRow(34);
            HSSFCell rac1 = rack.getCell(4);
            HSSFCell rac2 = rack.getCell(5);
            HSSFCell rac3 = rack.getCell(3);
            c76.setText(rac1.toString());
            c77.setText(rac2.toString());
            c78.setText(rac3.toString());

            HSSFRow rack1 = COMMUNICATION.getRow(35);
            HSSFCell racr1 = rack1.getCell(4);
            HSSFCell racr2 = rack1.getCell(5);
            HSSFCell racr3 = rack1.getCell(3);
            c79.setText(racr1.toString());
            c80.setText(racr2.toString());
            c81.setText(racr3.toString());

            HSSFRow rack2 = COMMUNICATION.getRow(36);
            HSSFCell race1 = rack2.getCell(4);
            HSSFCell race2 = rack2.getCell(5);
            HSSFCell race3 = rack2.getCell(3);
            c82.setText(race1.toString());
            c83.setText(race2.toString());
            c84.setText(race3.toString());

            HSSFRow rackt = COMMUNICATION.getRow(37);
            HSSFCell ract1 = rackt.getCell(4);
            HSSFCell ract2 = rackt.getCell(5);
            HSSFCell ract3 = rackt.getCell(3);
            c85.setText(ract1.toString());
            c86.setText(ract2.toString());
            c87.setText(ract3.toString());

            HSSFRow rackw = COMMUNICATION.getRow(38);
            HSSFCell racw1 = rackw.getCell(4);
            HSSFCell racw2 = rackw.getCell(5);
            HSSFCell racw3 = rackw.getCell(3);
            c88.setText(racw1.toString());
            c89.setText(racw2.toString());
            c90.setText(racw3.toString());

            HSSFRow rackf = COMMUNICATION.getRow(39);
            HSSFCell racf1 = rackf.getCell(4);
            HSSFCell racf2 = rackf.getCell(5);
            HSSFCell racf3 = rackf.getCell(3);
            c91.setText(racf1.toString());
            c92.setText(racf2.toString());
            c93.setText(racf3.toString());

            HSSFRow racks = COMMUNICATION.getRow(40);
            HSSFCell racs1 = racks.getCell(4);
            HSSFCell racs2 = racks.getCell(5);
            HSSFCell racs3 = racks.getCell(3);
            c94.setText(racs1.toString());
            c95.setText(racs2.toString());
            c96.setText(racs3.toString());
            HSSFRow rackk = COMMUNICATION.getRow(41);
            HSSFCell racu1 = rackk.getCell(4);
            HSSFCell racu2 = rackk.getCell(5);
            HSSFCell racu3 = rackk.getCell(3);
            c97.setText(racu1.toString());
            c98.setText(racu2.toString());
            c99.setText(racu3.toString());
            HSSFRow rackz = COMMUNICATION.getRow(42);
            HSSFCell racz1 = rackz.getCell(4);
            HSSFCell racz2 = rackz.getCell(5);
            HSSFCell racz3 = rackz.getCell(3);
            c100.setText(racz1.toString());
            c101.setText(racz2.toString());
            c102.setText(racz3.toString());
            HSSFRow racka = COMMUNICATION.getRow(43);
            HSSFCell raca1 = racka.getCell(4);
            HSSFCell raca2 = racka.getCell(5);
            HSSFCell raca3 = racka.getCell(3);
            c103.setText(raca1.toString());
            c104.setText(raca2.toString());
            c105.setText(raca3.toString());
            HSSFRow rackq = COMMUNICATION.getRow(44);
            HSSFCell racq1 = rackq.getCell(4);
            HSSFCell racq2 = rackq.getCell(5);
            HSSFCell racq3 = rackq.getCell(3);
            c106.setText(racq1.toString());
            c107.setText(racq2.toString());
            c108.setText(racq3.toString());
            HSSFRow racki = COMMUNICATION.getRow(45);
            HSSFCell raci1 = racki.getCell(4);
            HSSFCell raci2 = racki.getCell(5);
            HSSFCell raci3 = racki.getCell(3);
            c109.setText(raci1.toString());
            c110.setText(raci2.toString());
            c111.setText(raci3.toString());
            HSSFRow rackl = COMMUNICATION.getRow(46);
            HSSFCell racl1 = rackl.getCell(4);
            HSSFCell racl2 = rackl.getCell(5);
            HSSFCell racl3 = rackl.getCell(3);
            c112.setText(racl1.toString());
            c113.setText(racl2.toString());
            c114.setText(racl3.toString());
            HSSFRow rackm = COMMUNICATION.getRow(47);
            HSSFCell racm1 = rackm.getCell(4);
            HSSFCell racm2 = rackm.getCell(5);
            HSSFCell racm3 = rackm.getCell(3);
            c115.setText(racm1.toString());
            c116.setText(racm2.toString());
            c117.setText(racm3.toString());
            HSSFRow rackn = COMMUNICATION.getRow(48);
            HSSFCell racn1 = rackn.getCell(4);
            HSSFCell racn2 = rackn.getCell(5);
            HSSFCell racn3 = rackn.getCell(3);
            c118.setText(racn1.toString());
            c119.setText(racn2.toString());
            c120.setText(racn3.toString());
            HSSFRow rackb = COMMUNICATION.getRow(49);
            HSSFCell racb1 = rackb.getCell(4);
            HSSFCell racb2 = rackb.getCell(5);
            HSSFCell racb3 = rackb.getCell(3);
            c121.setText(racb1.toString());
            c122.setText(racb2.toString());
            c123.setText(racb3.toString());
            HSSFRow rackv = COMMUNICATION.getRow(50);
            HSSFCell racv1 = rackv.getCell(4);
            HSSFCell racv2 = rackv.getCell(5);
            HSSFCell racv3 = rackv.getCell(3);
            c124.setText(racv1.toString());
            c125.setText(racv2.toString());
            c126.setText(racv3.toString());
            HSSFRow rackg = COMMUNICATION.getRow(51);
            HSSFCell racg1 = rackg.getCell(4);
            HSSFCell racg2 = rackg.getCell(5);
            HSSFCell racg3 = rackg.getCell(3);
            c127.setText(racg1.toString());
            c128.setText(racg2.toString());
            c129.setText(racg3.toString());
            HSSFRow rackx = COMMUNICATION.getRow(52);
            HSSFCell racx1 = rackx.getCell(4);
            HSSFCell racx2 = rackx.getCell(5);
            HSSFCell racx3 = rackx.getCell(3);
            c130.setText(racx1.toString());
            c131.setText(racx2.toString());
            c132.setText(racx3.toString());
            HSSFRow rackj = COMMUNICATION.getRow(53);
            HSSFCell racj1 = rackj.getCell(4);
            HSSFCell racj2 = rackj.getCell(5);
            HSSFCell racj3 = rackj.getCell(3);
            c133.setText(racj1.toString());
            c134.setText(racj2.toString());
            c135.setText(racj3.toString());

        } catch (Exception e) {
            e.printStackTrace();
        }
    }
    private void SaveDatainxcel(Context context, String filename) {

        try {
            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);

            HSSFWorkbook workbook = new HSSFWorkbook(myInputStream);
            HSSFSheet CONTAINER = workbook.getSheetAt(11);

            HSSFRow cond = CONTAINER.getRow(5);
            HSSFCell con1 = cond.getCell(4);
            HSSFCell con2 = cond.getCell(5);
            HSSFCell con3 = cond.getCell(3);
            con1.setCellValue(c1.getText().toString().trim());
            con2.setCellValue(c2.getText().toString().trim());
            con3.setCellValue(c3.getText().toString().trim());

            HSSFRow evap = CONTAINER.getRow(6);
            HSSFCell eva1 = evap.getCell(4);
            HSSFCell eva2 = evap.getCell(5);
            HSSFCell eva3 = evap.getCell(3);
            eva1.setCellValue(c4.getText().toString().trim());
            eva2.setCellValue(c5.getText().toString().trim());
            eva3.setCellValue(c6.getText().toString().trim());

            HSSFRow comp = CONTAINER.getRow(7);
            HSSFCell com1 = comp.getCell(4);
            HSSFCell com2 = comp.getCell(5);
            HSSFCell com3 = comp.getCell(3);
            com1.setCellValue(c7.getText().toString().trim());
            com2.setCellValue(c8.getText().toString().trim());
            com3.setCellValue(c9.getText().toString().trim());

            HSSFRow mult = CONTAINER.getRow(8);
            HSSFCell mul1 = mult.getCell(4);
            HSSFCell mul2 = mult.getCell(5);
            HSSFCell mul3 = mult.getCell(3);
            mul1.setCellValue(c10.getText().toString().trim());
            mul2.setCellValue(c11.getText().toString().trim());
            mul3.setCellValue(c12.getText().toString().trim());

            HSSFRow ligt = CONTAINER.getRow(9);
            HSSFCell lig1 = ligt.getCell(4);
            HSSFCell lig2 = ligt.getCell(5);
            HSSFCell lig3 = ligt.getCell(3);
            lig1.setCellValue(c13.getText().toString().trim());
            lig2.setCellValue(c14.getText().toString().trim());
            lig3.setCellValue(c15.getText().toString().trim());

            HSSFRow term = CONTAINER.getRow(10);
            HSSFCell ter1 = term.getCell(4);
            HSSFCell ter2 = term.getCell(5);
            HSSFCell ter3 = term.getCell(3);
            ter1.setCellValue(c16.getText().toString().trim());
            ter2.setCellValue(c17.getText().toString().trim());
            ter3.setCellValue(c18.getText().toString().trim());

            HSSFRow fusi = CONTAINER.getRow(11);
            HSSFCell fus1 = fusi.getCell(4);
            HSSFCell fus2 = fusi.getCell(5);
            HSSFCell fus3 = fusi.getCell(3);
            fus1.setCellValue(c19.getText().toString().trim());
            fus2.setCellValue(c20.getText().toString().trim());
            fus3.setCellValue(c21.getText().toString().trim());

            HSSFRow freq = CONTAINER.getRow(12);
            HSSFCell fre1 = freq.getCell(4);
            HSSFCell fre2 = freq.getCell(5);
            HSSFCell fre3 = freq.getCell(3);
            fre1.setCellValue(c22.getText().toString().trim());
            fre2.setCellValue(c23.getText().toString().trim());
            fre3.setCellValue(c24.getText().toString().trim());

            HSSFRow powr = CONTAINER.getRow(13);
            HSSFCell pow1 = powr.getCell(4);
            HSSFCell pow2 = powr.getCell(5);
            HSSFCell pow3 = powr.getCell(3);
            pow1.setCellValue(c25.getText().toString().trim());
            pow2.setCellValue(c26.getText().toString().trim());
            pow3.setCellValue(c27.getText().toString().trim());

            HSSFRow ufan = CONTAINER.getRow(14);
            HSSFCell fan1 = ufan.getCell(4);
            HSSFCell fan2 = ufan.getCell(5);
            HSSFCell fan3 = ufan.getCell(3);
            fan1.setCellValue(c28.getText().toString().trim());
            fan2.setCellValue(c29.getText().toString().trim());
            fan3.setCellValue(c30.getText().toString().trim());

            HSSFRow upsi = CONTAINER.getRow(15);
            HSSFCell ups1 = upsi.getCell(4);
            HSSFCell ups2 = upsi.getCell(5);
            HSSFCell ups3 = upsi.getCell(3);
            ups1.setCellValue(c31.getText().toString().trim());
            ups2.setCellValue(c32.getText().toString().trim());
            ups3.setCellValue(c33.getText().toString().trim());

            HSSFRow purg = CONTAINER.getRow(16);
            HSSFCell pur1 = purg.getCell(4);
            HSSFCell pur2 = purg.getCell(5);
            HSSFCell pur3 = purg.getCell(3);
            pur1.setCellValue(c34.getText().toString().trim());
            pur2.setCellValue(c35.getText().toString().trim());
            pur3.setCellValue(c36.getText().toString().trim());

            HSSFRow smok = CONTAINER.getRow(17);
            HSSFCell smk1 = smok.getCell(4);
            HSSFCell smk2 = smok.getCell(5);
            HSSFCell smk3 = smok.getCell(3);
            smk1.setCellValue(c37.getText().toString().trim());
            smk2.setCellValue(c38.getText().toString().trim());
            smk3.setCellValue(c39.getText().toString().trim());

            HSSFRow fire = CONTAINER.getRow(18);
            HSSFCell fire1 = fire.getCell(4);
            HSSFCell fire2 = fire.getCell(5);
            HSSFCell fire3 = fire.getCell(3);
            fire1.setCellValue(c40.getText().toString().trim());
            fire2.setCellValue(c41.getText().toString().trim());
            fire3.setCellValue(c42.getText().toString().trim());

            HSSFRow toxi = CONTAINER.getRow(19);
            HSSFCell tox1 = toxi.getCell(4);
            HSSFCell tox2 = toxi.getCell(5);
            HSSFCell tox3 = toxi.getCell(3);
            tox1.setCellValue(c43.getText().toString().trim());
            tox2.setCellValue(c44.getText().toString().trim());
            tox3.setCellValue(c45.getText().toString().trim());

            HSSFRow flam = CONTAINER.getRow(20);
            HSSFCell flm1 = flam.getCell(4);
            HSSFCell flm2 = flam.getCell(5);
            HSSFCell flm3 = flam.getCell(3);
            flm1.setCellValue(c46.getText().toString().trim());
            flm2.setCellValue(c47.getText().toString().trim());
            flm3.setCellValue(c48.getText().toString().trim());

            HSSFRow faid = CONTAINER.getRow(21);
            HSSFCell fai1 = faid.getCell(4);
            HSSFCell fai2 = faid.getCell(5);
            HSSFCell fai3 = faid.getCell(3);
            fai1.setCellValue(c49.getText().toString().trim());
            fai2.setCellValue(c50.getText().toString().trim());
            fai3.setCellValue(c51.getText().toString().trim());

            HSSFRow radi = CONTAINER.getRow(22);
            HSSFCell rad1 = radi.getCell(4);
            HSSFCell rad2 = radi.getCell(5);
            HSSFCell rad3 = radi.getCell(3);
            rad1.setCellValue(c52.getText().toString().trim());
            rad2.setCellValue(c53.getText().toString().trim());
            rad3.setCellValue(c54.getText().toString().trim());

            HSSFRow refr = CONTAINER.getRow(23);
            HSSFCell ref1 = refr.getCell(4);
            HSSFCell ref2 = refr.getCell(5);
            HSSFCell ref3 = refr.getCell(3);
            ref1.setCellValue(c55.getText().toString().trim());
            ref2.setCellValue(c56.getText().toString().trim());
            ref3.setCellValue(c57.getText().toString().trim());

            HSSFRow sink = CONTAINER.getRow(24);
            HSSFCell sin1 = sink.getCell(4);
            HSSFCell sin2 = sink.getCell(5);
            HSSFCell sin3 = sink.getCell(3);
            sin1.setCellValue(c58.getText().toString().trim());
            sin2.setCellValue(c59.getText().toString().trim());
            sin3.setCellValue(c60.getText().toString().trim());

            HSSFRow exti = CONTAINER.getRow(25);
            HSSFCell ext1 = exti.getCell(4);
            HSSFCell ext2 = exti.getCell(5);
            HSSFCell ext3 = exti.getCell(3);
            ext1.setCellValue(c61.getText().toString().trim());
            ext2.setCellValue(c62.getText().toString().trim());
            ext3.setCellValue(c63.getText().toString().trim());

            HSSFRow chai = CONTAINER.getRow(26);
            HSSFCell cha1 = chai.getCell(4);
            HSSFCell cha2 = chai.getCell(5);
            HSSFCell cha3 = chai.getCell(3);
            cha1.setCellValue(c64.getText().toString().trim());
            cha2.setCellValue(c65.getText().toString().trim());
            cha3.setCellValue(c66.getText().toString().trim());

            HSSFRow thrm = CONTAINER.getRow(27);
            HSSFCell thr1 = thrm.getCell(4);
            HSSFCell thr2 = thrm.getCell(5);
            HSSFCell thr3 = thrm.getCell(3);
            thr1.setCellValue(c67.getText().toString().trim());
            thr2.setCellValue(c68.getText().toString().trim());
            thr3.setCellValue(c69.getText().toString().trim());

            HSSFRow h2mo = CONTAINER.getRow(28);
            HSSFCell h2m1 = h2mo.getCell(4);
            HSSFCell h2m2 = h2mo.getCell(5);
            HSSFCell h2m3 = h2mo.getCell(3);
            h2m1.setCellValue(c70.getText().toString().trim());
            h2m2.setCellValue(c71.getText().toString().trim());
            h2m3.setCellValue(c72.getText().toString().trim());

            HSSFRow antn = CONTAINER.getRow(29);
            HSSFCell ant1 = antn.getCell(4);
            HSSFCell ant2 = antn.getCell(5);
            HSSFCell ant3 = antn.getCell(3);
            ant1.setCellValue(c73.getText().toString().trim());
            ant2.setCellValue(c74.getText().toString().trim());
            ant3.setCellValue(c75.getText().toString().trim());

            HSSFRow rack = CONTAINER.getRow(34);
            HSSFCell rac1 = rack.getCell(4);
            HSSFCell rac2 = rack.getCell(5);
            HSSFCell rac3 = rack.getCell(3);
            rac1.setCellValue(c76.getText().toString().trim());
            rac2.setCellValue(c77.getText().toString().trim());
            rac3.setCellValue(c78.getText().toString().trim());

            HSSFRow rackq = CONTAINER.getRow(35);
            HSSFCell racq1 = rackq.getCell(4);
            HSSFCell racq2 = rackq.getCell(5);
            HSSFCell racq3 = rackq.getCell(3);
            racq1.setCellValue(c79.getText().toString().trim());
            racq2.setCellValue(c80.getText().toString().trim());
            racq3.setCellValue(c81.getText().toString().trim());

            HSSFRow rackw = CONTAINER.getRow(36);
            HSSFCell racw1 = rackw.getCell(4);
            HSSFCell racw2 = rackw.getCell(5);
            HSSFCell racw3 = rackw.getCell(3);
            racw1.setCellValue(c82.getText().toString().trim());
            racw2.setCellValue(c83.getText().toString().trim());
            racw3.setCellValue(c84.getText().toString().trim());

            HSSFRow racke = CONTAINER.getRow(37);
            HSSFCell race1 = racke.getCell(4);
            HSSFCell race2 = racke.getCell(5);
            HSSFCell race3 = racke.getCell(3);
            race1.setCellValue(c85.getText().toString().trim());
            race2.setCellValue(c86.getText().toString().trim());
            race3.setCellValue(c87.getText().toString().trim());

            HSSFRow rackr = CONTAINER.getRow(38);
            HSSFCell racr1 = rackr.getCell(4);
            HSSFCell racr2 = rackr.getCell(5);
            HSSFCell racr3 = rackr.getCell(3);
            racr1.setCellValue(c88.getText().toString().trim());
            racr2.setCellValue(c89.getText().toString().trim());
            racr3.setCellValue(c90.getText().toString().trim());

            HSSFRow rackt = CONTAINER.getRow(39);
            HSSFCell ract1 = rackt.getCell(4);
            HSSFCell ract2 = rackt.getCell(5);
            HSSFCell ract3 = rackt.getCell(3);
            ract1.setCellValue(c91.getText().toString().trim());
            ract2.setCellValue(c92.getText().toString().trim());
            ract3.setCellValue(c93.getText().toString().trim());

            HSSFRow racky = CONTAINER.getRow(40);
            HSSFCell racy1 = racky.getCell(4);
            HSSFCell racy2 = racky.getCell(5);
            HSSFCell racy3 = racky.getCell(3);
            racy1.setCellValue(c94.getText().toString().trim());
            racy2.setCellValue(c95.getText().toString().trim());
            racy3.setCellValue(c96.getText().toString().trim());

            HSSFRow racku = CONTAINER.getRow(41);
            HSSFCell racu1 = racku.getCell(4);
            HSSFCell racu2 = racku.getCell(5);
            HSSFCell racu3 = racku.getCell(3);
            racu1.setCellValue(c97.getText().toString().trim());
            racu2.setCellValue(c98.getText().toString().trim());
            racu3.setCellValue(c99.getText().toString().trim());

            HSSFRow racki = CONTAINER.getRow(42);
            HSSFCell raci1 = racki.getCell(4);
            HSSFCell raci2 = racki.getCell(5);
            HSSFCell raci3 = racki.getCell(3);
            raci1.setCellValue(c100.getText().toString().trim());
            raci2.setCellValue(c101.getText().toString().trim());
            raci3.setCellValue(c102.getText().toString().trim());

            HSSFRow racko = CONTAINER.getRow(43);
            HSSFCell raco1 = racko.getCell(4);
            HSSFCell raco2 = racko.getCell(5);
            HSSFCell raco3 = racko.getCell(3);
            raco1.setCellValue(c103.getText().toString().trim());
            raco2.setCellValue(c104.getText().toString().trim());
            raco3.setCellValue(c105.getText().toString().trim());

            HSSFRow rackp = CONTAINER.getRow(44);
            HSSFCell racp1 = rackp.getCell(4);
            HSSFCell racp2 = rackp.getCell(5);
            HSSFCell racp3 = rackp.getCell(3);
            racp1.setCellValue(c106.getText().toString().trim());
            racp2.setCellValue(c107.getText().toString().trim());
            racp3.setCellValue(c108.getText().toString().trim());

            HSSFRow racka = CONTAINER.getRow(45);
            HSSFCell raca1 = racka.getCell(4);
            HSSFCell raca2 = racka.getCell(5);
            HSSFCell raca3 = racka.getCell(3);
            raca1.setCellValue(c109.getText().toString().trim());
            raca2.setCellValue(c110.getText().toString().trim());
            raca3.setCellValue(c111.getText().toString().trim());

            HSSFRow racks = CONTAINER.getRow(46);
            HSSFCell racs1 = racks.getCell(4);
            HSSFCell racs2 = racks.getCell(5);
            HSSFCell racs3 = racks.getCell(3);
            racs1.setCellValue(c112.getText().toString().trim());
            racs2.setCellValue(c113.getText().toString().trim());
            racs3.setCellValue(c114.getText().toString().trim());

            HSSFRow rackd = CONTAINER.getRow(47);
            HSSFCell racd1 = rackd.getCell(4);
            HSSFCell racd2 = rackd.getCell(5);
            HSSFCell racd3 = rackd.getCell(3);
            racd1.setCellValue(c115.getText().toString().trim());
            racd2.setCellValue(c116.getText().toString().trim());
            racd3.setCellValue(c117.getText().toString().trim());

            HSSFRow rackf = CONTAINER.getRow(58);
            HSSFCell racf1 = rackf.getCell(4);
            HSSFCell racf2 = rackf.getCell(5);
            HSSFCell racf3 = rackf.getCell(3);
            racf1.setCellValue(c118.getText().toString().trim());
            racf2.setCellValue(c119.getText().toString().trim());
            racf3.setCellValue(c120.getText().toString().trim());

            HSSFRow rackg = CONTAINER.getRow(49);
            HSSFCell racg1 = rackg.getCell(4);
            HSSFCell racg2 = rackg.getCell(5);
            HSSFCell racg3 = rackg.getCell(3);
            racg1.setCellValue(c121.getText().toString().trim());
            racg2.setCellValue(c122.getText().toString().trim());
            racg3.setCellValue(c123.getText().toString().trim());

            HSSFRow rackh = CONTAINER.getRow(50);
            HSSFCell rach1 = rackh.getCell(4);
            HSSFCell rach2 = rackh.getCell(5);
            HSSFCell rach3 = rackh.getCell(3);
            rach1.setCellValue(c124.getText().toString().trim());
            rach2.setCellValue(c125.getText().toString().trim());
            rach3.setCellValue(c126.getText().toString().trim());

            HSSFRow rackj = CONTAINER.getRow(51);
            HSSFCell racj1 = rackj.getCell(4);
            HSSFCell racj2 = rackj.getCell(5);
            HSSFCell racj3 = rackj.getCell(3);
            racj1.setCellValue(c127.getText().toString().trim());
            racj2.setCellValue(c128.getText().toString().trim());
            racj3.setCellValue(c129.getText().toString().trim());

            HSSFRow rackk = CONTAINER.getRow(52);
            HSSFCell rack1 = rackk.getCell(4);
            HSSFCell rack2 = rackk.getCell(5);
            HSSFCell rack3 = rackk.getCell(3);
            rack1.setCellValue(c130.getText().toString().trim());
            rack2.setCellValue(c131.getText().toString().trim());
            rack3.setCellValue(c132.getText().toString().trim());

            HSSFRow rackl = CONTAINER.getRow(53);
            HSSFCell racl1 = rackl.getCell(4);
            HSSFCell racl2 = rackl.getCell(5);
            HSSFCell racl3 = rackl.getCell(3);
            racl1.setCellValue(c133.getText().toString().trim());
            racl2.setCellValue(c134.getText().toString().trim());
            racl3.setCellValue(c135.getText().toString().trim());

            myInputStream.close();
            FileOutputStream fos = new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename));
            workbook.write(fos);
            fos.close();
            Toast.makeText(context, "Changes are done.", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
        }}
    }

