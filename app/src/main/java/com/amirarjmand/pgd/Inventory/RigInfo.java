package com.amirarjmand.pgd.Inventory;

import android.content.Context;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.ScrollView;
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

public class RigInfo extends AppCompatActivity {
    EditText c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16, c17, c18, c19, c20, c21, c22, c23, c24, c25, c26, c27, c28, c29, c30, c31, c32, c33, c34, c35, c36, c37, c38, c39, c40, c41, c42, c43, c44, c45, c46, c47, c48, c49, c50, c51, c52, c53, c54, c55, c56, c57, c58, c59, c60, c61, c62, c63, c64,c65,c66, c67, c68, c69, c70, c71, c72, c73, c74, c75, c76, c77, c78, c79, c80, c81, c82, c83, c84, c85, c86, c87, c88, c89, c90, c91, c92, c93, c94, c95, c96, c97, c98, c99, c100, c101, c102, c103, c104, c105, c106, c107, c108, c109, c110, c111, c112, c113, c114, c115, c116, c117, c118, c119, c120, c121, c122, c123, c124, c125, c126, c127, c128, c129, c130, c131, c132, c133, c134, c135, c136, c137, c138, c139, c140, c141, c142, c143, c144, c145, c146, c147, c148, c149, c150, c151, c152, c153, c154, c155, c156, c157, c158, c159, c160, c161, c162, c163, c164, c165, c166, c167, c168, c169, c170, c171, c172, c173, c174, c175, c176, c177, c178, c179, c180, c181, c182, c183, c184, c185, c186, c187, c188, c189, c190, c191, c192, c193, c194, c195, c196, c197, c198, c199, c200, c201, c202, c203, c204, c205, c206, c207, c208, c209, c210, c211, c212, c213, c214, c215, c216, c217, c218, c219, c220, c221, c222, c223, c224, c225, c226, c227, c228, c229, c230, c231, c232, c233, c234, c235, c236, c237, c238, c239, c240, c241, c242, c243, c244, c245, c246, c247, c248, c249, c250, c251, c252, c253, c254, c255, c256, c257, c258, c259, c260, c261, c262, c263, c264, c265, c266, c267, c268, c269, c270, c271, c272, c273, c274, c275, c276, c277, c278, c279, c280, c281, c282, c283, c284, c285, c286, c287, c288, c289, c290, c291, c292, c293, c294;;

    ScrollView scrollView;
    Button save;
    String UnitNo;
    TextView header;
    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(RigInfo.this);
        setContentView(R.layout.activity_rig_info);



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
        c136 = findViewById(R.id.Co136);
        c137 = findViewById(R.id.Co137);
        c138 = findViewById(R.id.Co138);
        c139 = findViewById(R.id.Co139);
        c140 = findViewById(R.id.Co140);
        c141 = findViewById(R.id.Co141);
        c142 = findViewById(R.id.Co142);
        c143 = findViewById(R.id.Co143);
        c144 = findViewById(R.id.Co144);
        c145 = findViewById(R.id.Co145);
        c146 = findViewById(R.id.Co146);
        c147 = findViewById(R.id.Co147);
        c148 = findViewById(R.id.Co148);
        c149 = findViewById(R.id.Co149);
        c150 = findViewById(R.id.Co150);
        c151 = findViewById(R.id.Co151);
        c152 = findViewById(R.id.Co152);
        c153 = findViewById(R.id.Co153);
        c154 = findViewById(R.id.Co154);
        c155 = findViewById(R.id.Co155);
        c156 = findViewById(R.id.Co156);
        c157 = findViewById(R.id.Co157);
        c158 = findViewById(R.id.Co158);
        c159 = findViewById(R.id.Co159);
        c160 = findViewById(R.id.Co160);
        c161 = findViewById(R.id.Co161);
        c162 = findViewById(R.id.Co162);
        c163 = findViewById(R.id.Co163);
        c164 = findViewById(R.id.Co164);
        c165 = findViewById(R.id.Co165);
        c166 = findViewById(R.id.Co166);
        c167 = findViewById(R.id.Co167);
        c168 = findViewById(R.id.Co168);
        c169 = findViewById(R.id.Co169);
        c170 = findViewById(R.id.Co170);
        c171 = findViewById(R.id.Co171);
        c172 = findViewById(R.id.Co172);
        c173 = findViewById(R.id.Co173);
        c174 = findViewById(R.id.Co174);
        c175 = findViewById(R.id.Co175);
        c176 = findViewById(R.id.Co176);
        c177 = findViewById(R.id.Co177);
        c178 = findViewById(R.id.Co178);
        c179 = findViewById(R.id.Co179);
        c180 = findViewById(R.id.Co180);
        c181 = findViewById(R.id.Co181);
        c182 = findViewById(R.id.Co182);
        c183 = findViewById(R.id.Co183);
        c184 = findViewById(R.id.Co184);
        c185 = findViewById(R.id.Co185);
        c186 = findViewById(R.id.Co186);
        c187 = findViewById(R.id.Co187);
        c188 = findViewById(R.id.Co188);
        c189 = findViewById(R.id.Co189);
        c190 = findViewById(R.id.Co190);
        c191 = findViewById(R.id.Co191);
        c192 = findViewById(R.id.Co192);
        c193 = findViewById(R.id.Co193);
        c194 = findViewById(R.id.Co194);
        c195 = findViewById(R.id.Co195);
        c196 = findViewById(R.id.Co196);
        c197 = findViewById(R.id.Co197);
        c198 = findViewById(R.id.Co198);
        c199 = findViewById(R.id.Co199);
        c200 = findViewById(R.id.Co200);
        c201 = findViewById(R.id.Co201);
        c202 = findViewById(R.id.Co202);
        c203 = findViewById(R.id.Co203);
        c204 = findViewById(R.id.Co204);
        c205 = findViewById(R.id.Co205);
        c206 = findViewById(R.id.Co206);
        c207 = findViewById(R.id.Co207);
        c208 = findViewById(R.id.Co208);
        c209 = findViewById(R.id.Co209);
        c210 = findViewById(R.id.Co210);
        c211 = findViewById(R.id.Co211);
        c212 = findViewById(R.id.Co212);
        c213 = findViewById(R.id.Co213);
        c214 = findViewById(R.id.Co214);
        c215 = findViewById(R.id.Co215);
        c216 = findViewById(R.id.Co216);
        c217 = findViewById(R.id.Co217);
        c218 = findViewById(R.id.Co218);
        c219 = findViewById(R.id.Co219);
        c220 = findViewById(R.id.Co220);
        c221 = findViewById(R.id.Co221);
        c222 = findViewById(R.id.Co222);
        c223 = findViewById(R.id.Co223);
        c224 = findViewById(R.id.Co224);
        c225 = findViewById(R.id.Co225);
        c226 = findViewById(R.id.Co226);
        c227 = findViewById(R.id.Co227);
        c228 = findViewById(R.id.Co228);
        c229 = findViewById(R.id.Co229);
        c230 = findViewById(R.id.Co230);
        c231 = findViewById(R.id.Co231);
        c232 = findViewById(R.id.Co232);
        c233 = findViewById(R.id.Co233);
        c234 = findViewById(R.id.Co234);
        c235 = findViewById(R.id.Co235);
        c236 = findViewById(R.id.Co236);
        c237 = findViewById(R.id.Co237);
        c238 = findViewById(R.id.Co238);
        c239 = findViewById(R.id.Co239);
        c240 = findViewById(R.id.Co240);
        c241 = findViewById(R.id.Co241);
        c242 = findViewById(R.id.Co242);
        c243 = findViewById(R.id.Co243);
        c244 = findViewById(R.id.Co244);
        c245 = findViewById(R.id.Co245);
        c246 = findViewById(R.id.Co246);
        c247 = findViewById(R.id.Co247);
        c248 = findViewById(R.id.Co248);
        c249 = findViewById(R.id.Co249);
        c250 = findViewById(R.id.Co250);
        c251 = findViewById(R.id.Co251);
        c252 = findViewById(R.id.Co252);
        c253 = findViewById(R.id.Co253);
        c254 = findViewById(R.id.Co254);
        c255 = findViewById(R.id.Co255);
        c256 = findViewById(R.id.Co256);
        c257 = findViewById(R.id.Co257);
        c258 = findViewById(R.id.Co258);
        c259 = findViewById(R.id.Co259);
        c260 = findViewById(R.id.Co260);
        c261 = findViewById(R.id.Co261);
        c262 = findViewById(R.id.Co262);
        c263 = findViewById(R.id.Co263);
        c264 = findViewById(R.id.Co264);
        c265 = findViewById(R.id.Co265);
        c266 = findViewById(R.id.Co266);
        c267 = findViewById(R.id.Co267);
        c268 = findViewById(R.id.Co268);
        c269 = findViewById(R.id.Co269);
        c270 = findViewById(R.id.Co270);
        c271 = findViewById(R.id.Co271);
        c272 = findViewById(R.id.Co272);
        c273 = findViewById(R.id.Co273);
        c274 = findViewById(R.id.Co274);
        c275 = findViewById(R.id.Co275);
        c276 = findViewById(R.id.Co276);
        c277 = findViewById(R.id.Co277);
        c278 = findViewById(R.id.Co278);
        c279 = findViewById(R.id.Co279);
        c280 = findViewById(R.id.Co280);
        c281 = findViewById(R.id.Co281);
        c282 = findViewById(R.id.Co282);
        c283 = findViewById(R.id.Co283);
        c284 = findViewById(R.id.Co284);
        c285 = findViewById(R.id.Co285);
        c286 = findViewById(R.id.Co286);
        c287 = findViewById(R.id.Co287);
        c288 = findViewById(R.id.Co288);
        c289 = findViewById(R.id.Co289);
        c290 = findViewById(R.id.Co290);
        c291 = findViewById(R.id.Co291);
        c292 = findViewById(R.id.Co292);
        c293 = findViewById(R.id.Co293);
        c294 = findViewById(R.id.Co294);




        header=findViewById(R.id.HeaderC) ;

        Bundle Uno=getIntent().getExtras();
        UnitNo=Uno.getString("uno");
        header.setText("Rig Information System of unit "+UnitNo);
        scrollView=findViewById(R.id.Scr);

        save=findViewById(R.id.Save);

   


        ReadDataFromExcel(RigInfo.this, "Inventory List unit "+UnitNo+".xls");

        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                SaveDatainxcel(RigInfo.this, "Inventory List unit "+UnitNo+".xls");

            }
        });


    }
    private void ReadDataFromExcel(RigInfo RigInfo, String filename) {

        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook inv = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet COMMUNICATION = inv.getSheetAt(2);

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

            HSSFRow fusi = COMMUNICATION.getRow(12);
            HSSFCell fus1 = fusi.getCell(4);
            HSSFCell fus2 = fusi.getCell(5);
            HSSFCell fus3 = fusi.getCell(3);
            c19.setText(fus1.toString());
            c20.setText(fus2.toString());
            c21.setText(fus3.toString());

            HSSFRow powr = COMMUNICATION.getRow(13);
            HSSFCell pow1 = powr.getCell(4);
            HSSFCell pow2 = powr.getCell(5);
            HSSFCell pow3 = powr.getCell(3);
            c22.setText(pow1.toString());
            c23.setText(pow2.toString());
            c24.setText(pow3.toString());

            HSSFRow trns = COMMUNICATION.getRow(14);
            HSSFCell trn1 = trns.getCell(4);
            HSSFCell trn2 = trns.getCell(5);
            HSSFCell trn3 = trns.getCell(3);
            c25.setText(trn1.toString());
            c26.setText(trn2.toString());
            c27.setText(trn3.toString());

            HSSFRow fani = COMMUNICATION.getRow(15);
            HSSFCell fan1 = fani.getCell(4);
            HSSFCell fan2 = fani.getCell(5);
            HSSFCell fan3 = fani.getCell(3);
            c28.setText(fan1.toString());
            c29.setText(fan2.toString());
            c30.setText(fan3.toString());

            HSSFRow upsi = COMMUNICATION.getRow(16);
            HSSFCell ups1 = upsi.getCell(4);
            HSSFCell ups2 = upsi.getCell(5);
            HSSFCell ups3 = upsi.getCell(3);
            c31.setText(ups1.toString());
            c32.setText(ups2.toString());
            c33.setText(ups3.toString());

            HSSFRow purg = COMMUNICATION.getRow(19);
            HSSFCell pur1 = purg.getCell(4);
            HSSFCell pur2 = purg.getCell(5);
            HSSFCell pur3 = purg.getCell(3);
            c34.setText(pur1.toString());
            c35.setText(pur2.toString());
            c36.setText(pur3.toString());

            HSSFRow smok = COMMUNICATION.getRow(20);
            HSSFCell smo1 = smok.getCell(4);
            HSSFCell smo2 = smok.getCell(5);
            HSSFCell smo3 = smok.getCell(3);
            c37.setText(smo1.toString());
            c38.setText(smo2.toString());
            c39.setText(smo3.toString());

            HSSFRow fire = COMMUNICATION.getRow(21);
            HSSFCell fir1 = fire.getCell(4);
            HSSFCell fir2 = fire.getCell(5);
            HSSFCell fir3 = fire.getCell(3);
            c40.setText(fir1.toString());
            c41.setText(fir2.toString());
            c42.setText(fir3.toString());

            HSSFRow toxi = COMMUNICATION.getRow(22);
            HSSFCell tox1 = toxi.getCell(4);
            HSSFCell tox2 = toxi.getCell(5);
            HSSFCell tox3 = toxi.getCell(3);
            c43.setText(tox1.toString());
            c44.setText(tox2.toString());
            c45.setText(tox3.toString());

            HSSFRow flam = COMMUNICATION.getRow(23);
            HSSFCell flm1 = flam.getCell(4);
            HSSFCell flm2 = flam.getCell(5);
            HSSFCell flm3 = flam.getCell(3);
            c46.setText(flm1.toString());
            c47.setText(flm2.toString());
            c48.setText(flm3.toString());

            HSSFRow firs = COMMUNICATION.getRow(24);
            HSSFCell fira1 = firs.getCell(4);
            HSSFCell fira2 = firs.getCell(5);
            HSSFCell fira3 = firs.getCell(3);
            c49.setText(fira1.toString());
            c50.setText(fira2.toString());
            c51.setText(fira3.toString());

            HSSFRow radi = COMMUNICATION.getRow(35);
            HSSFCell rad1 = radi.getCell(4);
            HSSFCell rad2 = radi.getCell(5);
            HSSFCell rad3 = radi.getCell(3);
            c52.setText(rad1.toString());
            c53.setText(rad2.toString());
            c54.setText(rad3.toString());

            HSSFRow refg = COMMUNICATION.getRow(36);
            HSSFCell ref1 = refg.getCell(4);
            HSSFCell ref2 = refg.getCell(5);
            HSSFCell ref3 = refg.getCell(3);
            c55.setText(ref1.toString());
            c56.setText(ref2.toString());
            c57.setText(ref3.toString());

            HSSFRow sink = COMMUNICATION.getRow(37);
            HSSFCell sin1 = sink.getCell(4);
            HSSFCell sin2 = sink.getCell(5);
            HSSFCell sin3 = sink.getCell(3);
            c58.setText(sin1.toString());
            c59.setText(sin2.toString());
            c60.setText(sin3.toString());

            HSSFRow exti = COMMUNICATION.getRow(38);
            HSSFCell ext1 = exti.getCell(4);
            HSSFCell ext2 = exti.getCell(5);
            HSSFCell ext3 = exti.getCell(3);
            c61.setText(ext1.toString());
            c62.setText(ext2.toString());
            c63.setText(ext3.toString());

            HSSFRow chai = COMMUNICATION.getRow(39);
            HSSFCell cha1 = chai.getCell(4);
            HSSFCell cha2 = chai.getCell(5);
            HSSFCell cha3 = chai.getCell(3);
            c64.setText(cha1.toString());
            c65.setText(cha2.toString());
            c66.setText(cha3.toString());

            HSSFRow thrm = COMMUNICATION.getRow(41);
            HSSFCell thr1 = thrm.getCell(4);
            HSSFCell thr2 = thrm.getCell(5);
            HSSFCell thr3 = thrm.getCell(3);
            c67.setText(thr1.toString());
            c68.setText(thr2.toString());
            c69.setText(thr3.toString());

            HSSFRow hydr = COMMUNICATION.getRow(42);
            HSSFCell hyd1 = hydr.getCell(4);
            HSSFCell hyd2 = hydr.getCell(5);
            HSSFCell hyd3 = hydr.getCell(3);
            c70.setText(hyd1.toString());
            c71.setText(hyd2.toString());
            c72.setText(hyd3.toString());

            HSSFRow antn = COMMUNICATION.getRow(43);
            HSSFCell ant1 = antn.getCell(4);
            HSSFCell ant2 = antn.getCell(5);
            HSSFCell ant3 = antn.getCell(3);
            c73.setText(ant1.toString());
            c74.setText(ant2.toString());
            c75.setText(ant3.toString());

            HSSFRow rack = COMMUNICATION.getRow(44);
            HSSFCell rac1 = rack.getCell(4);
            HSSFCell rac2 = rack.getCell(5);
            HSSFCell rac3 = rack.getCell(3);
            c76.setText(rac1.toString());
            c77.setText(rac2.toString());
            c78.setText(rac3.toString());

            HSSFRow rack1 = COMMUNICATION.getRow(45);
            HSSFCell racr1 = rack1.getCell(4);
            HSSFCell racr2 = rack1.getCell(5);
            HSSFCell racr3 = rack1.getCell(3);
            c79.setText(racr1.toString());
            c80.setText(racr2.toString());
            c81.setText(racr3.toString());

            HSSFRow rack2 = COMMUNICATION.getRow(48);
            HSSFCell race1 = rack2.getCell(4);
            HSSFCell race2 = rack2.getCell(5);
            HSSFCell race3 = rack2.getCell(3);
            c82.setText(race1.toString());
            c83.setText(race2.toString());
            c84.setText(race3.toString());

            HSSFRow rackt = COMMUNICATION.getRow(49);
            HSSFCell ract1 = rackt.getCell(4);
            HSSFCell ract2 = rackt.getCell(5);
            HSSFCell ract3 = rackt.getCell(3);
            c85.setText(ract1.toString());
            c86.setText(ract2.toString());
            c87.setText(ract3.toString());

            HSSFRow rackw = COMMUNICATION.getRow(50);
            HSSFCell racw1 = rackw.getCell(4);
            HSSFCell racw2 = rackw.getCell(5);
            HSSFCell racw3 = rackw.getCell(3);
            c88.setText(racw1.toString());
            c89.setText(racw2.toString());
            c90.setText(racw3.toString());

            HSSFRow rackf = COMMUNICATION.getRow(51);
            HSSFCell racf1 = rackf.getCell(4);
            HSSFCell racf2 = rackf.getCell(5);
            HSSFCell racf3 = rackf.getCell(3);
            c91.setText(racf1.toString());
            c92.setText(racf2.toString());
            c93.setText(racf3.toString());

            HSSFRow racks = COMMUNICATION.getRow(52);
            HSSFCell racs1 = racks.getCell(4);
            HSSFCell racs2 = racks.getCell(5);
            HSSFCell racs3 = racks.getCell(3);
            c94.setText(racs1.toString());
            c95.setText(racs2.toString());
            c96.setText(racs3.toString());
            HSSFRow rackk = COMMUNICATION.getRow(53);
            HSSFCell racu1 = rackk.getCell(4);
            HSSFCell racu2 = rackk.getCell(5);
            HSSFCell racu3 = rackk.getCell(3);
            c97.setText(racu1.toString());
            c98.setText(racu2.toString());
            c99.setText(racu3.toString());
            HSSFRow rackz = COMMUNICATION.getRow(54);
            HSSFCell racz1 = rackz.getCell(4);
            HSSFCell racz2 = rackz.getCell(5);
            HSSFCell racz3 = rackz.getCell(3);
            c100.setText(racz1.toString());
            c101.setText(racz2.toString());
            c102.setText(racz3.toString());
            HSSFRow racka = COMMUNICATION.getRow(56);
            HSSFCell raca1 = racka.getCell(4);
            HSSFCell raca2 = racka.getCell(5);
            HSSFCell raca3 = racka.getCell(3);
            c103.setText(raca1.toString());
            c104.setText(raca2.toString());
            c105.setText(raca3.toString());
            HSSFRow rackq = COMMUNICATION.getRow(57);
            HSSFCell racq1 = rackq.getCell(4);
            HSSFCell racq2 = rackq.getCell(5);
            HSSFCell racq3 = rackq.getCell(3);
            c106.setText(racq1.toString());
            c107.setText(racq2.toString());
            c108.setText(racq3.toString());
            HSSFRow racki = COMMUNICATION.getRow(58);
            HSSFCell raci1 = racki.getCell(4);
            HSSFCell raci2 = racki.getCell(5);
            HSSFCell raci3 = racki.getCell(3);
            c109.setText(raci1.toString());
            c110.setText(raci2.toString());
            c111.setText(raci3.toString());
            HSSFRow rackl = COMMUNICATION.getRow(59);
            HSSFCell racl1 = rackl.getCell(4);
            HSSFCell racl2 = rackl.getCell(5);
            HSSFCell racl3 = rackl.getCell(3);
            c112.setText(racl1.toString());
            c113.setText(racl2.toString());
            c114.setText(racl3.toString());
            HSSFRow rackm = COMMUNICATION.getRow(60);
            HSSFCell racm1 = rackm.getCell(4);
            HSSFCell racm2 = rackm.getCell(5);
            HSSFCell racm3 = rackm.getCell(3);
            c115.setText(racm1.toString());
            c116.setText(racm2.toString());
            c117.setText(racm3.toString());
            HSSFRow rackn = COMMUNICATION.getRow(61);
            HSSFCell racn1 = rackn.getCell(4);
            HSSFCell racn2 = rackn.getCell(5);
            HSSFCell racn3 = rackn.getCell(3);
            c118.setText(racn1.toString());
            c119.setText(racn2.toString());
            c120.setText(racn3.toString());
            HSSFRow rackb = COMMUNICATION.getRow(62);
            HSSFCell racb1 = rackb.getCell(4);
            HSSFCell racb2 = rackb.getCell(5);
            HSSFCell racb3 = rackb.getCell(3);
            c121.setText(racb1.toString());
            c122.setText(racb2.toString());
            c123.setText(racb3.toString());
            HSSFRow rackv = COMMUNICATION.getRow(69);
            HSSFCell racv1 = rackv.getCell(4);
            HSSFCell racv2 = rackv.getCell(5);
            HSSFCell racv3 = rackv.getCell(3);
            c124.setText(racv1.toString());
            c125.setText(racv2.toString());
            c126.setText(racv3.toString());
            HSSFRow rackg = COMMUNICATION.getRow(70);
            HSSFCell racg1 = rackg.getCell(4);
            HSSFCell racg2 = rackg.getCell(5);
            HSSFCell racg3 = rackg.getCell(3);
            c127.setText(racg1.toString());
            c128.setText(racg2.toString());
            c129.setText(racg3.toString());
            HSSFRow rackx = COMMUNICATION.getRow(71);
            HSSFCell racx1 = rackx.getCell(4);
            HSSFCell racx2 = rackx.getCell(5);
            HSSFCell racx3 = rackx.getCell(3);
            c130.setText(racx1.toString());
            c131.setText(racx2.toString());
            c132.setText(racx3.toString());
            HSSFRow rackj = COMMUNICATION.getRow(73);
            HSSFCell racj1 = rackj.getCell(4);
            HSSFCell racj2 = rackj.getCell(5);
            HSSFCell racj3 = rackj.getCell(3);
            c133.setText(racj1.toString());
            c134.setText(racj2.toString());
            c135.setText(racj3.toString());
            HSSFRow racko = COMMUNICATION.getRow(74);
            HSSFCell raco1 = racko.getCell(4);
            HSSFCell raco2 = racko.getCell(5);
            HSSFCell raco3 = racko.getCell(3);
            c136.setText(raco1.toString());
            c137.setText(raco2.toString());
            c138.setText(raco3.toString());
            HSSFRow rackp = COMMUNICATION.getRow(75);
            HSSFCell racp1 = rackp.getCell(4);
            HSSFCell racp2 = rackp.getCell(5);
            HSSFCell racp3 = rackp.getCell(3);
            c139.setText(racp1.toString());
            c140.setText(racp2.toString());
            c141.setText(racp3.toString());
            HSSFRow Condenserq = COMMUNICATION.getRow(77);
            HSSFCell ConCelql = Condenserq.getCell(4);
            HSSFCell ConCelq2 = Condenserq.getCell(5);
            HSSFCell ConCelq3 = Condenserq.getCell(3);
            c142.setText(ConCelql.toString());
            c143.setText(ConCelq2.toString());
            c144.setText(ConCelq3.toString());

            HSSFRow Evaq = COMMUNICATION.getRow(78);
            HSSFCell evaql = Evaq.getCell(4);
            HSSFCell evaq2 = Evaq.getCell(5);
            HSSFCell evaq3 = Evaq.getCell(3);
            c145.setText(evaql.toString());
            c146.setText(evaq2.toString());
            c147.setText(evaq3.toString());

            HSSFRow compq = COMMUNICATION.getRow(79);
            HSSFCell comq1 = compq.getCell(4);
            HSSFCell comq2 = compq.getCell(5);
            HSSFCell comq3 = compq.getCell(3);
            c148.setText(comq1.toString());
            c149.setText(comq2.toString());
            c150.setText(comq3.toString());

            HSSFRow multq = COMMUNICATION.getRow(81);
            HSSFCell mulq1 = multq.getCell(4);
            HSSFCell mulq2 = multq.getCell(5);
            HSSFCell mulq3 = multq.getCell(3);
            c151.setText(mulq1.toString());
            c152.setText(mulq2.toString());
            c153.setText(mulq3.toString());

            HSSFRow lighq = COMMUNICATION.getRow(82);
            HSSFCell ligq1 = lighq.getCell(4);
            HSSFCell ligq2 = lighq.getCell(5);
            HSSFCell ligq3 = lighq.getCell(3);
            c154.setText(ligq1.toString());
            c155.setText(ligq2.toString());
            c156.setText(ligq3.toString());

            HSSFRow termq = COMMUNICATION.getRow(83);
            HSSFCell terq1 = termq.getCell(4);
            HSSFCell terq2 = termq.getCell(5);
            HSSFCell terq3 = termq.getCell(3);
            c157.setText(terq1.toString());
            c158.setText(terq2.toString());
            c159.setText(terq3.toString());

            HSSFRow fusiq = COMMUNICATION.getRow(84);
            HSSFCell fusq1 = fusiq.getCell(4);
            HSSFCell fusq2 = fusiq.getCell(5);
            HSSFCell fusq3 = fusiq.getCell(3);
            c160.setText(fusq1.toString());
            c161.setText(fusq2.toString());
            c162.setText(fusq3.toString());

            HSSFRow powrq = COMMUNICATION.getRow(87);
            HSSFCell powq1 = powrq.getCell(4);
            HSSFCell powq2 = powrq.getCell(5);
            HSSFCell powq3 = powrq.getCell(3);
            c163.setText(powq1.toString());
            c164.setText(powq2.toString());
            c165.setText(powq3.toString());

            HSSFRow trnsz = COMMUNICATION.getRow(88);
            HSSFCell trnz1 = trnsz.getCell(4);
            HSSFCell trnz2 = trnsz.getCell(5);
            HSSFCell trnz3 = trnsz.getCell(3);
            c166.setText(trnz1.toString());
            c167.setText(trnz2.toString());
            c168.setText(trnz3.toString());

            HSSFRow faniz = COMMUNICATION.getRow(89);
            HSSFCell fanz1 = faniz.getCell(4);
            HSSFCell fanz2 = faniz.getCell(5);
            HSSFCell fanz3 = faniz.getCell(3);
            c169.setText(fanz1.toString());
            c170.setText(fanz2.toString());
            c171.setText(fanz3.toString());

            HSSFRow upsiz = COMMUNICATION.getRow(90);
            HSSFCell upsz1 = upsiz.getCell(4);
            HSSFCell upsz2 = upsiz.getCell(5);
            HSSFCell upsz3 = upsiz.getCell(3);
            c172.setText(upsz1.toString());
            c173.setText(upsz2.toString());
            c174.setText(upsz3.toString());

            HSSFRow purgz = COMMUNICATION.getRow(91);
            HSSFCell purz1 = purgz.getCell(4);
            HSSFCell purz2 = purgz.getCell(5);
            HSSFCell purz3 = purgz.getCell(3);
            c175.setText(purz1.toString());
            c176.setText(purz2.toString());
            c177.setText(purz3.toString());

            HSSFRow smokz = COMMUNICATION.getRow(92);
            HSSFCell smoz1 = smokz.getCell(4);
            HSSFCell smoz2 = smokz.getCell(5);
            HSSFCell smoz3 = smokz.getCell(3);
            c178.setText(smoz1.toString());
            c179.setText(smoz2.toString());
            c180.setText(smoz3.toString());

            HSSFRow firez = COMMUNICATION.getRow(93);
            HSSFCell firz1 = firez.getCell(4);
            HSSFCell firz2 = firez.getCell(5);
            HSSFCell firz3 = firez.getCell(3);
            c181.setText(firz1.toString());
            c182.setText(firz2.toString());
            c183.setText(firz3.toString());

            HSSFRow toxiz = COMMUNICATION.getRow(94);
            HSSFCell toxz1 = toxiz.getCell(4);
            HSSFCell toxz2 = toxiz.getCell(5);
            HSSFCell toxz3 = toxiz.getCell(3);
            c184.setText(toxz1.toString());
            c185.setText(toxz2.toString());
            c186.setText(toxz3.toString());

            HSSFRow flamz = COMMUNICATION.getRow(95);
            HSSFCell flmz1 = flamz.getCell(4);
            HSSFCell flmz2 = flamz.getCell(5);
            HSSFCell flmz3 = flamz.getCell(3);
            c187.setText(flmz1.toString());
            c188.setText(flmz2.toString());
            c189.setText(flmz3.toString());

            HSSFRow firsz = COMMUNICATION.getRow(102);
            HSSFCell firaz1 = firsz.getCell(4);
            HSSFCell firaz2 = firsz.getCell(5);
            HSSFCell firaz3 = firsz.getCell(3);
            c190.setText(firaz1.toString());
            c191.setText(firaz2.toString());
            c192.setText(firaz3.toString());

            HSSFRow radiz = COMMUNICATION.getRow(103);
            HSSFCell radz1 = radiz.getCell(4);
            HSSFCell radz2 = radiz.getCell(5);
            HSSFCell radz3 = radiz.getCell(3);
            c193.setText(radz1.toString());
            c194.setText(radz2.toString());
            c195.setText(radz3.toString());

            HSSFRow refgz = COMMUNICATION.getRow(104);
            HSSFCell refz1 = refgz.getCell(4);
            HSSFCell refz2 = refgz.getCell(5);
            HSSFCell refz3 = refgz.getCell(3);
            c196.setText(refz1.toString());
            c197.setText(refz2.toString());
            c198.setText(refz3.toString());

            HSSFRow sinkz = COMMUNICATION.getRow(105);
            HSSFCell sinz1 = sinkz.getCell(4);
            HSSFCell sinz2 = sinkz.getCell(5);
            HSSFCell sinz3 = sinkz.getCell(3);
            c199.setText(sinz1.toString());
            c200.setText(sinz2.toString());
            c201.setText(sinz3.toString());

            HSSFRow extiz = COMMUNICATION.getRow(106);
            HSSFCell extz1 = extiz.getCell(4);
            HSSFCell extz2 = extiz.getCell(5);
            HSSFCell extz3 = extiz.getCell(3);
            c202.setText(extz1.toString());
            c203.setText(extz2.toString());
            c204.setText(extz3.toString());

            HSSFRow chaiz = COMMUNICATION.getRow(107);
            HSSFCell chaz1 = chaiz.getCell(4);
            HSSFCell chaz2 = chaiz.getCell(5);
            HSSFCell chaz3 = chaiz.getCell(3);
            c205.setText(chaz1.toString());
            c206.setText(chaz2.toString());
            c207.setText(chaz3.toString());

            HSSFRow thrmc = COMMUNICATION.getRow(108);
            HSSFCell thrc1 = thrmc.getCell(4);
            HSSFCell thrc2 = thrmc.getCell(5);
            HSSFCell thrc3 = thrmc.getCell(3);
            c208.setText(thrc1.toString());
            c209.setText(thrc2.toString());
            c210.setText(thrc3.toString());

            HSSFRow hydro = COMMUNICATION.getRow(109);
            HSSFCell hydo1 = hydro.getCell(4);
            HSSFCell hydo2 = hydro.getCell(5);
            HSSFCell hydo3 = hydro.getCell(3);
            c211.setText(hydo1.toString());
            c212.setText(hydo2.toString());
            c213.setText(hydo3.toString());

            HSSFRow antnp = COMMUNICATION.getRow(110);
            HSSFCell antp1 = antnp.getCell(4);
            HSSFCell antp2 = antnp.getCell(5);
            HSSFCell antp3 = antnp.getCell(3);
            c214.setText(antp1.toString());
            c215.setText(antp2.toString());
            c216.setText(antp3.toString());

            HSSFRow racktt = COMMUNICATION.getRow(111);
            HSSFCell racrt1 = racktt.getCell(4);
            HSSFCell racrt2 = racktt.getCell(5);
            HSSFCell racrt3 = racktt.getCell(3);
            c217.setText(racrt1.toString());
            c218.setText(racrt2.toString());
            c219.setText(racrt3.toString());
            HSSFRow Condenserw = COMMUNICATION.getRow(112);
            HSSFCell ConCelwl = Condenserw.getCell(4);
            HSSFCell ConCelw2 = Condenserw.getCell(5);
            HSSFCell ConCelw3 = Condenserw.getCell(3);
            c220.setText(ConCelwl.toString());
            c221.setText(ConCelw2.toString());
            c222.setText(ConCelw3.toString());

            HSSFRow Evaw = COMMUNICATION.getRow(113);
            HSSFCell evawl = Evaw.getCell(4);
            HSSFCell evaw2 = Evaw.getCell(5);
            HSSFCell evaw3 = Evaw.getCell(3);
            c223.setText(evawl.toString());
            c224.setText(evaw2.toString());
            c225.setText(evaw3.toString());

            HSSFRow compw = COMMUNICATION.getRow(114);
            HSSFCell comw1 = compw.getCell(4);
            HSSFCell comw2 = compw.getCell(5);
            HSSFCell comw3 = compw.getCell(3);
            c226.setText(comw1.toString());
            c227.setText(comw2.toString());
            c228.setText(comw3.toString());

            HSSFRow multw = COMMUNICATION.getRow(115);
            HSSFCell mulw1 = multw.getCell(4);
            HSSFCell mulw2 = multw.getCell(5);
            HSSFCell mulw3 = multw.getCell(3);
            c229.setText(mulw1.toString());
            c230.setText(mulw2.toString());
            c231.setText(mulw3.toString());

            HSSFRow lighw = COMMUNICATION.getRow(116);
            HSSFCell ligw1 = lighw.getCell(4);
            HSSFCell ligw2 = lighw.getCell(5);
            HSSFCell ligw3 = lighw.getCell(3);
            c232.setText(ligw1.toString());
            c233.setText(ligw2.toString());
            c234.setText(ligw3.toString());

            HSSFRow termw = COMMUNICATION.getRow(117);
            HSSFCell terw1 = termw.getCell(4);
            HSSFCell terw2 = termw.getCell(5);
            HSSFCell terw3 = termw.getCell(3);
            c235.setText(terw1.toString());
            c236.setText(terw2.toString());
            c237.setText(terw3.toString());

            HSSFRow fusiw = COMMUNICATION.getRow(118);
            HSSFCell fusw1 = fusiw.getCell(4);
            HSSFCell fusw2 = fusiw.getCell(5);
            HSSFCell fusw3 = fusiw.getCell(3);
            c238.setText(fusw1.toString());
            c239.setText(fusw2.toString());
            c240.setText(fusw3.toString());

            HSSFRow powrqq = COMMUNICATION.getRow(119);
            HSSFCell powqq1 = powrqq.getCell(4);
            HSSFCell powqq2 = powrqq.getCell(5);
            HSSFCell powqq3 = powrqq.getCell(3);
            c241.setText(powqq1.toString());
            c242.setText(powqq2.toString());
            c243.setText(powqq3.toString());

            HSSFRow trnsq = COMMUNICATION.getRow(120);
            HSSFCell trnq1 = trnsq.getCell(4);
            HSSFCell trnq2 = trnsq.getCell(5);
            HSSFCell trnq3 = trnsq.getCell(3);
            c244.setText(trnq1.toString());
            c245.setText(trnq2.toString());
            c246.setText(trnq3.toString());

            HSSFRow faniq = COMMUNICATION.getRow(121);
            HSSFCell fanq1 = faniq.getCell(4);
            HSSFCell fanq2 = faniq.getCell(5);
            HSSFCell fanq3 = faniq.getCell(3);
            c247.setText(fanq1.toString());
            c248.setText(fanq2.toString());
            c249.setText(fanq3.toString());

            HSSFRow upsiq = COMMUNICATION.getRow(122);
            HSSFCell upsq1 = upsiq.getCell(4);
            HSSFCell upsq2 = upsiq.getCell(5);
            HSSFCell upsq3 = upsiq.getCell(3);
            c250.setText(upsq1.toString());
            c251.setText(upsq2.toString());
            c252.setText(upsq3.toString());

            HSSFRow purgq = COMMUNICATION.getRow(123);
            HSSFCell purq1 = purgq.getCell(4);
            HSSFCell purq2 = purgq.getCell(5);
            HSSFCell purq3 = purgq.getCell(3);
            c253.setText(purq1.toString());
            c254.setText(purq2.toString());
            c255.setText(purq3.toString());

            HSSFRow smokq = COMMUNICATION.getRow(124);
            HSSFCell smoq1 = smokq.getCell(4);
            HSSFCell smoq2 = smokq.getCell(5);
            HSSFCell smoq3 = smokq.getCell(3);
            c256.setText(smoq1.toString());
            c257.setText(smoq2.toString());
            c258.setText(smoq3.toString());

            HSSFRow fireq = COMMUNICATION.getRow(125);
            HSSFCell firq1 = fireq.getCell(4);
            HSSFCell firq2 = fireq.getCell(5);
            HSSFCell firq3 = fireq.getCell(3);
            c259.setText(firq1.toString());
            c260.setText(firq2.toString());
            c261.setText(firq3.toString());

            HSSFRow toxiq = COMMUNICATION.getRow(126);
            HSSFCell toxq1 = toxiq.getCell(4);
            HSSFCell toxq2 = toxiq.getCell(5);
            HSSFCell toxq3 = toxiq.getCell(3);
            c262.setText(toxq1.toString());
            c263.setText(toxq2.toString());
            c264.setText(toxq3.toString());

            HSSFRow flamq = COMMUNICATION.getRow(127);
            HSSFCell flmq1 = flamq.getCell(4);
            HSSFCell flmq2 = flamq.getCell(5);
            HSSFCell flmq3 = flamq.getCell(3);
            c265.setText(flmq1.toString());
            c266.setText(flmq2.toString());
            c267.setText(flmq3.toString());

            HSSFRow firsq = COMMUNICATION.getRow(128);
            HSSFCell firaq1 = firsq.getCell(4);
            HSSFCell firaq2 = firsq.getCell(5);
            HSSFCell firaq3 = firsq.getCell(3);
            c268.setText(firaq1.toString());
            c269.setText(firaq2.toString());
            c270.setText(firaq3.toString());

            HSSFRow radiq = COMMUNICATION.getRow(134);
            HSSFCell radq1 = radiq.getCell(4);
            HSSFCell radq2 = radiq.getCell(5);
            HSSFCell radq3 = radiq.getCell(3);
            c271.setText(radq1.toString());
            c272.setText(radq2.toString());
            c273.setText(radq3.toString());

            HSSFRow refgq = COMMUNICATION.getRow(135);
            HSSFCell refq1 = refgq.getCell(4);
            HSSFCell refq2 = refgq.getCell(5);
            HSSFCell refq3 = refgq.getCell(3);
            c274.setText(refq1.toString());
            c275.setText(refq2.toString());
            c276.setText(refq3.toString());

            HSSFRow sinkq = COMMUNICATION.getRow(136);
            HSSFCell sinq1 = sinkq.getCell(4);
            HSSFCell sinq2 = sinkq.getCell(5);
            HSSFCell sinq3 = sinkq.getCell(3);
            c277.setText(sinq1.toString());
            c278.setText(sinq2.toString());
            c279.setText(sinq3.toString());

            HSSFRow extiq = COMMUNICATION.getRow(137);
            HSSFCell extq1 = extiq.getCell(4);
            HSSFCell extq2 = extiq.getCell(5);
            HSSFCell extq3 = extiq.getCell(3);
            c280.setText(extq1.toString());
            c281.setText(extq2.toString());
            c282.setText(extq3.toString());

            HSSFRow chaiq = COMMUNICATION.getRow(138);
            HSSFCell chaq1 = chaiq.getCell(4);
            HSSFCell chaq2 = chaiq.getCell(5);
            HSSFCell chaq3 = chaiq.getCell(3);
            c283.setText(chaq1.toString());
            c284.setText(chaq2.toString());
            c285.setText(chaq3.toString());

            HSSFRow thrmq = COMMUNICATION.getRow(139);
            HSSFCell thrq1 = thrmq.getCell(4);
            HSSFCell thrq2 = thrmq.getCell(5);
            HSSFCell thrq3 = thrmq.getCell(3);
            c286.setText(thrq1.toString());
            c287.setText(thrq2.toString());
            c288.setText(thrq3.toString());

            HSSFRow hydrq = COMMUNICATION.getRow(140);
            HSSFCell hydq1 = hydrq.getCell(4);
            HSSFCell hydq2 = hydrq.getCell(5);
            HSSFCell hydq3 = hydrq.getCell(3);
            c289.setText(hydq1.toString());
            c290.setText(hydq2.toString());
            c291.setText(hydq3.toString());

            HSSFRow antnq = COMMUNICATION.getRow(141);
            HSSFCell antq1 = antnq.getCell(4);
            HSSFCell antq2 = antnq.getCell(5);
            HSSFCell antq3 = antnq.getCell(3);
            c292.setText(antq1.toString());
            c293.setText(antq2.toString());
            c294.setText(antq3.toString());





        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void SaveDatainxcel(Context context, String filename) {

        try {
            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename);
            FileInputStream myInputStream = new FileInputStream(file);

            HSSFWorkbook workbook = new HSSFWorkbook (myInputStream);
            HSSFSheet CONTAINER = workbook.getSheetAt(2);

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

            HSSFRow fusi = CONTAINER.getRow(12);
            HSSFCell fus1 = fusi.getCell(4);
            HSSFCell fus2 = fusi.getCell(5);
            HSSFCell fus3 = fusi.getCell(3);
            fus1.setCellValue(c19.getText().toString().trim());
            fus2.setCellValue(c20.getText().toString().trim());
            fus3.setCellValue(c21.getText().toString().trim());

            HSSFRow freq = CONTAINER.getRow(13);
            HSSFCell fre1 = freq.getCell(4);
            HSSFCell fre2 = freq.getCell(5);
            HSSFCell fre3 = freq.getCell(3);
            fre1.setCellValue(c22.getText().toString().trim());
            fre2.setCellValue(c23.getText().toString().trim());
            fre3.setCellValue(c24.getText().toString().trim());

            HSSFRow powr = CONTAINER.getRow(14);
            HSSFCell pow1 = powr.getCell(4);
            HSSFCell pow2 = powr.getCell(5);
            HSSFCell pow3 = powr.getCell(3);
            pow1.setCellValue(c25.getText().toString().trim());
            pow2.setCellValue(c26.getText().toString().trim());
            pow3.setCellValue(c27.getText().toString().trim());

            HSSFRow ufan = CONTAINER.getRow(15);
            HSSFCell fan1 = ufan.getCell(4);
            HSSFCell fan2 = ufan.getCell(5);
            HSSFCell fan3 = ufan.getCell(3);
            fan1.setCellValue(c28.getText().toString().trim());
            fan2.setCellValue(c29.getText().toString().trim());
            fan3.setCellValue(c30.getText().toString().trim());

            HSSFRow upsi = CONTAINER.getRow(16);
            HSSFCell ups1 = upsi.getCell(4);
            HSSFCell ups2 = upsi.getCell(5);
            HSSFCell ups3 = upsi.getCell(3);
            ups1.setCellValue(c31.getText().toString().trim());
            ups2.setCellValue(c32.getText().toString().trim());
            ups3.setCellValue(c33.getText().toString().trim());

            HSSFRow purg = CONTAINER.getRow(19);
            HSSFCell pur1 = purg.getCell(4);
            HSSFCell pur2 = purg.getCell(5);
            HSSFCell pur3 = purg.getCell(3);
            pur1.setCellValue(c34.getText().toString().trim());
            pur2.setCellValue(c35.getText().toString().trim());
            pur3.setCellValue(c36.getText().toString().trim());

            HSSFRow smok = CONTAINER.getRow(20);
            HSSFCell smk1 = smok.getCell(4);
            HSSFCell smk2 = smok.getCell(5);
            HSSFCell smk3 = smok.getCell(3);
            smk1.setCellValue(c37.getText().toString().trim());
            smk2.setCellValue(c38.getText().toString().trim());
            smk3.setCellValue(c39.getText().toString().trim());

            HSSFRow fire = CONTAINER.getRow(21);
            HSSFCell fire1 = fire.getCell(4);
            HSSFCell fire2 = fire.getCell(5);
            HSSFCell fire3 = fire.getCell(3);
            fire1.setCellValue(c40.getText().toString().trim());
            fire2.setCellValue(c41.getText().toString().trim());
            fire3.setCellValue(c42.getText().toString().trim());

            HSSFRow toxi = CONTAINER.getRow(22);
            HSSFCell tox1 = toxi.getCell(4);
            HSSFCell tox2 = toxi.getCell(5);
            HSSFCell tox3 = toxi.getCell(3);
            tox1.setCellValue(c43.getText().toString().trim());
            tox2.setCellValue(c44.getText().toString().trim());
            tox3.setCellValue(c45.getText().toString().trim());

            HSSFRow flam = CONTAINER.getRow(23);
            HSSFCell flm1 = flam.getCell(4);
            HSSFCell flm2 = flam.getCell(5);
            HSSFCell flm3 = flam.getCell(3);
            flm1.setCellValue(c46.getText().toString().trim());
            flm2.setCellValue(c47.getText().toString().trim());
            flm3.setCellValue(c48.getText().toString().trim());

            HSSFRow faid = CONTAINER.getRow(24);
            HSSFCell fai1 = faid.getCell(4);
            HSSFCell fai2 = faid.getCell(5);
            HSSFCell fai3 = faid.getCell(3);
            fai1.setCellValue(c49.getText().toString().trim());
            fai2.setCellValue(c50.getText().toString().trim());
            fai3.setCellValue(c51.getText().toString().trim());

            HSSFRow radi = CONTAINER.getRow(35);
            HSSFCell rad1 = radi.getCell(4);
            HSSFCell rad2 = radi.getCell(5);
            HSSFCell rad3 = radi.getCell(3);
            rad1.setCellValue(c52.getText().toString().trim());
            rad2.setCellValue(c53.getText().toString().trim());
            rad3.setCellValue(c54.getText().toString().trim());

            HSSFRow refr = CONTAINER.getRow(36);
            HSSFCell ref1 = refr.getCell(4);
            HSSFCell ref2 = refr.getCell(5);
            HSSFCell ref3 = refr.getCell(3);
            ref1.setCellValue(c55.getText().toString().trim());
            ref2.setCellValue(c56.getText().toString().trim());
            ref3.setCellValue(c57.getText().toString().trim());

            HSSFRow sink = CONTAINER.getRow(37);
            HSSFCell sin1 = sink.getCell(4);
            HSSFCell sin2 = sink.getCell(5);
            HSSFCell sin3 = sink.getCell(3);
            sin1.setCellValue(c58.getText().toString().trim());
            sin2.setCellValue(c59.getText().toString().trim());
            sin3.setCellValue(c60.getText().toString().trim());

            HSSFRow exti = CONTAINER.getRow(38);
            HSSFCell ext1 = exti.getCell(4);
            HSSFCell ext2 = exti.getCell(5);
            HSSFCell ext3 = exti.getCell(3);
            ext1.setCellValue(c61.getText().toString().trim());
            ext2.setCellValue(c62.getText().toString().trim());
            ext3.setCellValue(c63.getText().toString().trim());

            HSSFRow chai = CONTAINER.getRow(39);
            HSSFCell cha1 = chai.getCell(4);
            HSSFCell cha2 = chai.getCell(5);
            HSSFCell cha3 = chai.getCell(3);
            cha1.setCellValue(c64.getText().toString().trim());
            cha2.setCellValue(c65.getText().toString().trim());
            cha3.setCellValue(c66.getText().toString().trim());

            HSSFRow thrm = CONTAINER.getRow(41);
            HSSFCell thr1 = thrm.getCell(4);
            HSSFCell thr2 = thrm.getCell(5);
            HSSFCell thr3 = thrm.getCell(3);
            thr1.setCellValue(c67.getText().toString().trim());
            thr2.setCellValue(c68.getText().toString().trim());
            thr3.setCellValue(c69.getText().toString().trim());

            HSSFRow h2mo = CONTAINER.getRow(42);
            HSSFCell h2m1 = h2mo.getCell(4);
            HSSFCell h2m2 = h2mo.getCell(5);
            HSSFCell h2m3 = h2mo.getCell(3);
            h2m1.setCellValue(c70.getText().toString().trim());
            h2m2.setCellValue(c71.getText().toString().trim());
            h2m3.setCellValue(c72.getText().toString().trim());

            HSSFRow antn = CONTAINER.getRow(43);
            HSSFCell ant1 = antn.getCell(4);
            HSSFCell ant2 = antn.getCell(5);
            HSSFCell ant3 = antn.getCell(3);
            ant1.setCellValue(c73.getText().toString().trim());
            ant2.setCellValue(c74.getText().toString().trim());
            ant3.setCellValue(c75.getText().toString().trim());

            HSSFRow rack = CONTAINER.getRow(44);
            HSSFCell rac1 = rack.getCell(4);
            HSSFCell rac2 = rack.getCell(5);
            HSSFCell rac3 = rack.getCell(3);
            rac1.setCellValue(c76.getText().toString().trim());
            rac2.setCellValue(c77.getText().toString().trim());
            rac3.setCellValue(c78.getText().toString().trim());

            HSSFRow rackq = CONTAINER.getRow(45);
            HSSFCell racq1 = rackq.getCell(4);
            HSSFCell racq2 = rackq.getCell(5);
            HSSFCell racq3 = rackq.getCell(3);
            racq1.setCellValue(c79.getText().toString().trim());
            racq2.setCellValue(c80.getText().toString().trim());
            racq3.setCellValue(c81.getText().toString().trim());

            HSSFRow rackw = CONTAINER.getRow(48);
            HSSFCell racw1 = rackw.getCell(4);
            HSSFCell racw2 = rackw.getCell(5);
            HSSFCell racw3 = rackw.getCell(3);
            racw1.setCellValue(c82.getText().toString().trim());
            racw2.setCellValue(c83.getText().toString().trim());
            racw3.setCellValue(c84.getText().toString().trim());

            HSSFRow racke = CONTAINER.getRow(49);
            HSSFCell race1 = racke.getCell(4);
            HSSFCell race2 = racke.getCell(5);
            HSSFCell race3 = racke.getCell(3);
            race1.setCellValue(c85.getText().toString().trim());
            race2.setCellValue(c86.getText().toString().trim());
            race3.setCellValue(c87.getText().toString().trim());

            HSSFRow rackr = CONTAINER.getRow(50);
            HSSFCell racr1 = rackr.getCell(4);
            HSSFCell racr2 = rackr.getCell(5);
            HSSFCell racr3 = rackr.getCell(3);
            racr1.setCellValue(c88.getText().toString().trim());
            racr2.setCellValue(c89.getText().toString().trim());
            racr3.setCellValue(c90.getText().toString().trim());

            HSSFRow rackt = CONTAINER.getRow(51);
            HSSFCell ract1 = rackt.getCell(4);
            HSSFCell ract2 = rackt.getCell(5);
            HSSFCell ract3 = rackt.getCell(3);
            ract1.setCellValue(c91.getText().toString().trim());
            ract2.setCellValue(c92.getText().toString().trim());
            ract3.setCellValue(c93.getText().toString().trim());

            HSSFRow racky = CONTAINER.getRow(52);
            HSSFCell racy1 = racky.getCell(4);
            HSSFCell racy2 = racky.getCell(5);
            HSSFCell racy3 = racky.getCell(3);
            racy1.setCellValue(c94.getText().toString().trim());
            racy2.setCellValue(c95.getText().toString().trim());
            racy3.setCellValue(c96.getText().toString().trim());

            HSSFRow racku = CONTAINER.getRow(53);
            HSSFCell racu1 = racku.getCell(4);
            HSSFCell racu2 = racku.getCell(5);
            HSSFCell racu3 = racku.getCell(3);
            racu1.setCellValue(c97.getText().toString().trim());
            racu2.setCellValue(c98.getText().toString().trim());
            racu3.setCellValue(c99.getText().toString().trim());

            HSSFRow racki = CONTAINER.getRow(54);
            HSSFCell raci1 = racki.getCell(4);
            HSSFCell raci2 = racki.getCell(5);
            HSSFCell raci3 = racki.getCell(3);
            raci1.setCellValue(c100.getText().toString().trim());
            raci2.setCellValue(c101.getText().toString().trim());
            raci3.setCellValue(c102.getText().toString().trim());

            HSSFRow racko = CONTAINER.getRow(56);
            HSSFCell raco1 = racko.getCell(4);
            HSSFCell raco2 = racko.getCell(5);
            HSSFCell raco3 = racko.getCell(3);
            raco1.setCellValue(c103.getText().toString().trim());
            raco2.setCellValue(c104.getText().toString().trim());
            raco3.setCellValue(c105.getText().toString().trim());

            HSSFRow rackp = CONTAINER.getRow(57);
            HSSFCell racp1 = rackp.getCell(4);
            HSSFCell racp2 = rackp.getCell(5);
            HSSFCell racp3 = rackp.getCell(3);
            racp1.setCellValue(c106.getText().toString().trim());
            racp2.setCellValue(c107.getText().toString().trim());
            racp3.setCellValue(c108.getText().toString().trim());

            HSSFRow racka = CONTAINER.getRow(58);
            HSSFCell raca1 = racka.getCell(4);
            HSSFCell raca2 = racka.getCell(5);
            HSSFCell raca3 = racka.getCell(3);
            raca1.setCellValue(c109.getText().toString().trim());
            raca2.setCellValue(c110.getText().toString().trim());
            raca3.setCellValue(c111.getText().toString().trim());

            HSSFRow racks = CONTAINER.getRow(59);
            HSSFCell racs1 = racks.getCell(4);
            HSSFCell racs2 = racks.getCell(5);
            HSSFCell racs3 = racks.getCell(3);
            racs1.setCellValue(c112.getText().toString().trim());
            racs2.setCellValue(c113.getText().toString().trim());
            racs3.setCellValue(c114.getText().toString().trim());

            HSSFRow rackd = CONTAINER.getRow(60);
            HSSFCell racd1 = rackd.getCell(4);
            HSSFCell racd2 = rackd.getCell(5);
            HSSFCell racd3 = rackd.getCell(3);
            racd1.setCellValue(c115.getText().toString().trim());
            racd2.setCellValue(c116.getText().toString().trim());
            racd3.setCellValue(c117.getText().toString().trim());

            HSSFRow rackf = CONTAINER.getRow(61);
            HSSFCell racf1 = rackf.getCell(4);
            HSSFCell racf2 = rackf.getCell(5);
            HSSFCell racf3 = rackf.getCell(3);
            racf1.setCellValue(c118.getText().toString().trim());
            racf2.setCellValue(c119.getText().toString().trim());
            racf3.setCellValue(c120.getText().toString().trim());

            HSSFRow rackg = CONTAINER.getRow(62);
            HSSFCell racg1 = rackg.getCell(4);
            HSSFCell racg2 = rackg.getCell(5);
            HSSFCell racg3 = rackg.getCell(3);
            racg1.setCellValue(c121.getText().toString().trim());
            racg2.setCellValue(c122.getText().toString().trim());
            racg3.setCellValue(c123.getText().toString().trim());

            HSSFRow rackh = CONTAINER.getRow(69);
            HSSFCell rach1 = rackh.getCell(4);
            HSSFCell rach2 = rackh.getCell(5);
            HSSFCell rach3 = rackh.getCell(3);
            rach1.setCellValue(c124.getText().toString().trim());
            rach2.setCellValue(c125.getText().toString().trim());
            rach3.setCellValue(c126.getText().toString().trim());

            HSSFRow rackj = CONTAINER.getRow(70);
            HSSFCell racj1 = rackj.getCell(4);
            HSSFCell racj2 = rackj.getCell(5);
            HSSFCell racj3 = rackj.getCell(3);
            racj1.setCellValue(c127.getText().toString().trim());
            racj2.setCellValue(c128.getText().toString().trim());
            racj3.setCellValue(c129.getText().toString().trim());

            HSSFRow rackk = CONTAINER.getRow(71);
            HSSFCell rack1 = rackk.getCell(4);
            HSSFCell rack2 = rackk.getCell(5);
            HSSFCell rack3 = rackk.getCell(3);
            rack1.setCellValue(c130.getText().toString().trim());
            rack2.setCellValue(c131.getText().toString().trim());
            rack3.setCellValue(c132.getText().toString().trim());

            HSSFRow rackl = CONTAINER.getRow(73);
            HSSFCell racl1 = rackl.getCell(4);
            HSSFCell racl2 = rackl.getCell(5);
            HSSFCell racl3 = rackl.getCell(3);
            racl1.setCellValue(c133.getText().toString().trim());
            racl2.setCellValue(c134.getText().toString().trim());
            racl3.setCellValue(c135.getText().toString().trim());

            HSSFRow rackz = CONTAINER.getRow(74);
            HSSFCell racz1 = rackz.getCell(4);
            HSSFCell racz2 = rackz.getCell(5);
            HSSFCell racz3 = rackz.getCell(3);
            racz1.setCellValue(c136.getText().toString().trim());
            racz2.setCellValue(c137.getText().toString().trim());
            racz3.setCellValue(c138.getText().toString().trim());

            HSSFRow rackx = CONTAINER.getRow(75);
            HSSFCell racx1 = rackx.getCell(4);
            HSSFCell racx2 = rackx.getCell(5);
            HSSFCell racx3 = rackx.getCell(3);
            racx1.setCellValue(c139.getText().toString().trim());
            racx2.setCellValue(c140.getText().toString().trim());
            racx3.setCellValue(c141.getText().toString().trim());

            HSSFRow condx = CONTAINER.getRow(77);
            HSSFCell conx1 = condx.getCell(4);
            HSSFCell conx2 = condx.getCell(5);
            HSSFCell conx3 = condx.getCell(3);
            conx1.setCellValue(c142.getText().toString().trim());
            conx2.setCellValue(c143.getText().toString().trim());
            conx3.setCellValue(c144.getText().toString().trim());

            HSSFRow evapx = CONTAINER.getRow(78);
            HSSFCell evax1 = evapx.getCell(4);
            HSSFCell evax2 = evapx.getCell(5);
            HSSFCell evax3 = evapx.getCell(3);
            evax1.setCellValue(c145.getText().toString().trim());
            evax2.setCellValue(c146.getText().toString().trim());
            evax3.setCellValue(c147.getText().toString().trim());

            HSSFRow compx = CONTAINER.getRow(79);
            HSSFCell comx1 = compx.getCell(4);
            HSSFCell comx2 = compx.getCell(5);
            HSSFCell comx3 = compx.getCell(3);
            comx1.setCellValue(c148.getText().toString().trim());
            comx2.setCellValue(c149.getText().toString().trim());
            comx3.setCellValue(c150.getText().toString().trim());

            HSSFRow multx = CONTAINER.getRow(81);
            HSSFCell mulx1 = multx.getCell(4);
            HSSFCell mulx2 = multx.getCell(5);
            HSSFCell mulx3 = multx.getCell(3);
            mulx1.setCellValue(c151.getText().toString().trim());
            mulx2.setCellValue(c152.getText().toString().trim());
            mulx3.setCellValue(c153.getText().toString().trim());

            HSSFRow ligtz = CONTAINER.getRow(82);
            HSSFCell ligz1 = ligtz.getCell(4);
            HSSFCell ligz2 = ligtz.getCell(5);
            HSSFCell ligz3 = ligtz.getCell(3);
            ligz1.setCellValue(c154.getText().toString().trim());
            ligz2.setCellValue(c155.getText().toString().trim());
            ligz3.setCellValue(c156.getText().toString().trim());

            HSSFRow termd = CONTAINER.getRow(83);
            HSSFCell terd1 = termd.getCell(4);
            HSSFCell terd2 = termd.getCell(5);
            HSSFCell terd3 = termd.getCell(3);
            terd1.setCellValue(c157.getText().toString().trim());
            terd2.setCellValue(c158.getText().toString().trim());
            terd3.setCellValue(c159.getText().toString().trim());

            HSSFRow fusir = CONTAINER.getRow(84);
            HSSFCell fusr1 = fusir.getCell(4);
            HSSFCell fusr2 = fusir.getCell(5);
            HSSFCell fusr3 = fusir.getCell(3);
            fusr1.setCellValue(c160.getText().toString().trim());
            fusr2.setCellValue(c161.getText().toString().trim());
            fusr3.setCellValue(c162.getText().toString().trim());

            HSSFRow freqr = CONTAINER.getRow(87);
            HSSFCell frer1 = freqr.getCell(4);
            HSSFCell frer2 = freqr.getCell(5);
            HSSFCell frer3 = freqr.getCell(3);
            frer1.setCellValue(c163.getText().toString().trim());
            frer2.setCellValue(c164.getText().toString().trim());
            frer3.setCellValue(c165.getText().toString().trim());

            HSSFRow powrm = CONTAINER.getRow(88);
            HSSFCell powm1 = powrm.getCell(4);
            HSSFCell powm2 = powrm.getCell(5);
            HSSFCell powm3 = powrm.getCell(3);
            powm1.setCellValue(c166.getText().toString().trim());
            powm2.setCellValue(c167.getText().toString().trim());
            powm3.setCellValue(c168.getText().toString().trim());

            HSSFRow ufann = CONTAINER.getRow(89);
            HSSFCell fann1 = ufann.getCell(4);
            HSSFCell fann2 = ufann.getCell(5);
            HSSFCell fann3 = ufann.getCell(3);
            fann1.setCellValue(c169.getText().toString().trim());
            fann2.setCellValue(c170.getText().toString().trim());
            fann3.setCellValue(c171.getText().toString().trim());

            HSSFRow upsin = CONTAINER.getRow(90);
            HSSFCell upsn1 = upsin.getCell(4);
            HSSFCell upsn2 = upsin.getCell(5);
            HSSFCell upsn3 = upsin.getCell(3);
            upsn1.setCellValue(c172.getText().toString().trim());
            upsn2.setCellValue(c173.getText().toString().trim());
            upsn3.setCellValue(c174.getText().toString().trim());

            HSSFRow purgn = CONTAINER.getRow(91);
            HSSFCell purn1 = purgn.getCell(4);
            HSSFCell purn2 = purgn.getCell(5);
            HSSFCell purn3 = purgn.getCell(3);
            purn1.setCellValue(c175.getText().toString().trim());
            purn2.setCellValue(c176.getText().toString().trim());
            purn3.setCellValue(c177.getText().toString().trim());

            HSSFRow smokg = CONTAINER.getRow(92);
            HSSFCell smkg1 = smokg.getCell(4);
            HSSFCell smkg2 = smokg.getCell(5);
            HSSFCell smkg3 = smokg.getCell(3);
            smkg1.setCellValue(c178.getText().toString().trim());
            smkg2.setCellValue(c179.getText().toString().trim());
            smkg3.setCellValue(c180.getText().toString().trim());

            HSSFRow firee = CONTAINER.getRow(93);
            HSSFCell firee1 = firee.getCell(4);
            HSSFCell firee2 = firee.getCell(5);
            HSSFCell firee3 = firee.getCell(3);
            firee1.setCellValue(c181.getText().toString().trim());
            firee2.setCellValue(c182.getText().toString().trim());
            firee3.setCellValue(c183.getText().toString().trim());

            HSSFRow toxib = CONTAINER.getRow(94);
            HSSFCell toxb1 = toxib.getCell(4);
            HSSFCell toxb2 = toxib.getCell(5);
            HSSFCell toxb3 = toxib.getCell(3);
            toxb1.setCellValue(c184.getText().toString().trim());
            toxb2.setCellValue(c185.getText().toString().trim());
            toxb3.setCellValue(c186.getText().toString().trim());

            HSSFRow flamb = CONTAINER.getRow(95);
            HSSFCell flmb1 = flamb.getCell(4);
            HSSFCell flmb2 = flamb.getCell(5);
            HSSFCell flmb3 = flamb.getCell(3);
            flmb1.setCellValue(c187.getText().toString().trim());
            flmb2.setCellValue(c188.getText().toString().trim());
            flmb3.setCellValue(c189.getText().toString().trim());

            HSSFRow faidq = CONTAINER.getRow(102);
            HSSFCell faiq1 = faidq.getCell(4);
            HSSFCell faiq2 = faidq.getCell(5);
            HSSFCell faiq3 = faidq.getCell(3);
            faiq1.setCellValue(c190.getText().toString().trim());
            faiq2.setCellValue(c191.getText().toString().trim());
            faiq3.setCellValue(c192.getText().toString().trim());

            HSSFRow radiq = CONTAINER.getRow(103);
            HSSFCell radq1 = radiq.getCell(4);
            HSSFCell radq2 = radiq.getCell(5);
            HSSFCell radq3 = radiq.getCell(3);
            radq1.setCellValue(c193.getText().toString().trim());
            radq2.setCellValue(c194.getText().toString().trim());
            radq3.setCellValue(c195.getText().toString().trim());

            HSSFRow refrg = CONTAINER.getRow(104);
            HSSFCell refg1 = refrg.getCell(4);
            HSSFCell refg2 = refrg.getCell(5);
            HSSFCell refg3 = refrg.getCell(3);
            refg1.setCellValue(c196.getText().toString().trim());
            refg2.setCellValue(c197.getText().toString().trim());
            refg3.setCellValue(c198.getText().toString().trim());

            HSSFRow sinkg = CONTAINER.getRow(105);
            HSSFCell sing1 = sinkg.getCell(4);
            HSSFCell sing2 = sinkg.getCell(5);
            HSSFCell sing3 = sinkg.getCell(3);
            sing1.setCellValue(c199.getText().toString().trim());
            sing2.setCellValue(c200.getText().toString().trim());
            sing3.setCellValue(c201.getText().toString().trim());

            HSSFRow extig = CONTAINER.getRow(106);
            HSSFCell extg1 = extig.getCell(4);
            HSSFCell extg2 = extig.getCell(5);
            HSSFCell extg3 = extig.getCell(3);
            extg1.setCellValue(c202.getText().toString().trim());
            extg2.setCellValue(c203.getText().toString().trim());
            extg3.setCellValue(c204.getText().toString().trim());

            HSSFRow chaif = CONTAINER.getRow(107);
            HSSFCell chaf1 = chaif.getCell(4);
            HSSFCell chaf2 = chaif.getCell(5);
            HSSFCell chaf3 = chaif.getCell(3);
            chaf1.setCellValue(c205.getText().toString().trim());
            chaf2.setCellValue(c206.getText().toString().trim());
            chaf3.setCellValue(c207.getText().toString().trim());

            HSSFRow thrmf = CONTAINER.getRow(108);
            HSSFCell thrf1 = thrmf.getCell(4);
            HSSFCell thrf2 = thrmf.getCell(5);
            HSSFCell thrf3 = thrmf.getCell(3);
            thrf1.setCellValue(c208.getText().toString().trim());
            thrf2.setCellValue(c209.getText().toString().trim());
            thrf3.setCellValue(c210.getText().toString().trim());

            HSSFRow h2mof = CONTAINER.getRow(109);
            HSSFCell h2mf1 = h2mof.getCell(4);
            HSSFCell h2mf2 = h2mof.getCell(5);
            HSSFCell h2mf3 = h2mof.getCell(3);
            h2mf1.setCellValue(c211.getText().toString().trim());
            h2mf2.setCellValue(c212.getText().toString().trim());
            h2mf3.setCellValue(c213.getText().toString().trim());

            HSSFRow antnf = CONTAINER.getRow(110);
            HSSFCell antf1 = antnf.getCell(4);
            HSSFCell antf2 = antnf.getCell(5);
            HSSFCell antf3 = antnf.getCell(3);
            antf1.setCellValue(c214.getText().toString().trim());
            antf2.setCellValue(c215.getText().toString().trim());
            antf3.setCellValue(c216.getText().toString().trim());

            HSSFRow rackzz = CONTAINER.getRow(111);
            HSSFCell raczz1 = rackzz.getCell(4);
            HSSFCell raczz2 = rackzz.getCell(5);
            HSSFCell raczz3 = rackzz.getCell(3);
            raczz1.setCellValue(c217.getText().toString().trim());
            raczz2.setCellValue(c218.getText().toString().trim());
            raczz3.setCellValue(c219.getText().toString().trim());

            HSSFRow rackqy = CONTAINER.getRow(112);
            HSSFCell racqy1 = rackqy.getCell(4);
            HSSFCell racqy2 = rackqy.getCell(5);
            HSSFCell racqy3 = rackqy.getCell(3);
            racqy1.setCellValue(c220.getText().toString().trim());
            racqy2.setCellValue(c221.getText().toString().trim());
            racqy3.setCellValue(c222.getText().toString().trim());

            HSSFRow rackwy = CONTAINER.getRow(113);
            HSSFCell racwy1 = rackwy.getCell(4);
            HSSFCell racwy2 = rackwy.getCell(5);
            HSSFCell racwy3 = rackwy.getCell(3);
            racwy1.setCellValue(c223.getText().toString().trim());
            racwy2.setCellValue(c224.getText().toString().trim());
            racwy3.setCellValue(c225.getText().toString().trim());

            HSSFRow rackey = CONTAINER.getRow(114);
            HSSFCell racey1 = rackey.getCell(4);
            HSSFCell racey2 = rackey.getCell(5);
            HSSFCell racey3 = rackey.getCell(3);
            racey1.setCellValue(c226.getText().toString().trim());
            racey2.setCellValue(c227.getText().toString().trim());
            racey3.setCellValue(c228.getText().toString().trim());

            HSSFRow rackrw = CONTAINER.getRow(115);
            HSSFCell racrw1 = rackrw.getCell(4);
            HSSFCell racrw2 = rackrw.getCell(5);
            HSSFCell racrw3 = rackrw.getCell(3);
            racrw1.setCellValue(c229.getText().toString().trim());
            racrw2.setCellValue(c230.getText().toString().trim());
            racrw3.setCellValue(c231.getText().toString().trim());

            HSSFRow rackts = CONTAINER.getRow(116);
            HSSFCell racts1 = rackts.getCell(4);
            HSSFCell racts2 = rackts.getCell(5);
            HSSFCell racts3 = rackts.getCell(3);
            racts1.setCellValue(c232.getText().toString().trim());
            racts2.setCellValue(c233.getText().toString().trim());
            racts3.setCellValue(c234.getText().toString().trim());

            HSSFRow rackys = CONTAINER.getRow(117);
            HSSFCell racys1 = rackys.getCell(4);
            HSSFCell racys2 = rackys.getCell(5);
            HSSFCell racys3 = rackys.getCell(3);
            racys1.setCellValue(c235.getText().toString().trim());
            racys2.setCellValue(c236.getText().toString().trim());
            racys3.setCellValue(c237.getText().toString().trim());

            HSSFRow rackus = CONTAINER.getRow(118);
            HSSFCell racus1 = rackus.getCell(4);
            HSSFCell racus2 = rackus.getCell(5);
            HSSFCell racus3 = rackus.getCell(3);
            racus1.setCellValue(c238.getText().toString().trim());
            racus2.setCellValue(c239.getText().toString().trim());
            racus3.setCellValue(c240.getText().toString().trim());

            HSSFRow rackis = CONTAINER.getRow(119);
            HSSFCell racis1 = rackis.getCell(4);
            HSSFCell racis2 = rackis.getCell(5);
            HSSFCell racis3 = rackis.getCell(3);
            racis1.setCellValue(c241.getText().toString().trim());
            racis2.setCellValue(c242.getText().toString().trim());
            racis3.setCellValue(c243.getText().toString().trim());

            HSSFRow rackos = CONTAINER.getRow(120);
            HSSFCell racos1 = rackos.getCell(4);
            HSSFCell racos2 = rackos.getCell(5);
            HSSFCell racos3 = rackos.getCell(3);
            racos1.setCellValue(c244.getText().toString().trim());
            racos2.setCellValue(c245.getText().toString().trim());
            racos3.setCellValue(c246.getText().toString().trim());

            HSSFRow rackps = CONTAINER.getRow(121);
            HSSFCell racps1 = rackps.getCell(4);
            HSSFCell racps2 = rackps.getCell(5);
            HSSFCell racps3 = rackps.getCell(3);
            racps1.setCellValue(c247.getText().toString().trim());
            racps2.setCellValue(c248.getText().toString().trim());
            racps3.setCellValue(c249.getText().toString().trim());

            HSSFRow rackas = CONTAINER.getRow(122);
            HSSFCell racas1 = rackas.getCell(4);
            HSSFCell racas2 = rackas.getCell(5);
            HSSFCell racas3 = rackas.getCell(3);
            racas1.setCellValue(c250.getText().toString().trim());
            racas2.setCellValue(c251.getText().toString().trim());
            racas3.setCellValue(c252.getText().toString().trim());

            HSSFRow rackss = CONTAINER.getRow(123);
            HSSFCell racss1 = rackss.getCell(4);
            HSSFCell racss2 = rackss.getCell(5);
            HSSFCell racss3 = rackss.getCell(3);
            racss1.setCellValue(c253.getText().toString().trim());
            racss2.setCellValue(c254.getText().toString().trim());
            racss3.setCellValue(c255.getText().toString().trim());

            HSSFRow rackda = CONTAINER.getRow(124);
            HSSFCell racda1 = rackda.getCell(4);
            HSSFCell racda2 = rackda.getCell(5);
            HSSFCell racda3 = rackda.getCell(3);
            racda1.setCellValue(c256.getText().toString().trim());
            racda2.setCellValue(c257.getText().toString().trim());
            racda3.setCellValue(c258.getText().toString().trim());

            HSSFRow rackaf = CONTAINER.getRow(125);
            HSSFCell racaf1 = rackaf.getCell(4);
            HSSFCell racaf2 = rackaf.getCell(5);
            HSSFCell racaf3 = rackaf.getCell(3);
            racaf1.setCellValue(c259.getText().toString().trim());
            racaf2.setCellValue(c260.getText().toString().trim());
            racaf3.setCellValue(c261.getText().toString().trim());

            HSSFRow rackaag = CONTAINER.getRow(126);
            HSSFCell racgaa1 = rackaag.getCell(4);
            HSSFCell racgaa2 = rackaag.getCell(5);
            HSSFCell racgaa3 = rackaag.getCell(3);
            racgaa1.setCellValue(c262.getText().toString().trim());
            racgaa2.setCellValue(c263.getText().toString().trim());
            racgaa3.setCellValue(c264.getText().toString().trim());

            HSSFRow rackqh = CONTAINER.getRow(127);
            HSSFCell rachq1 = rackqh.getCell(4);
            HSSFCell rachq2 = rackqh.getCell(5);
            HSSFCell rachq3 = rackqh.getCell(3);
            rachq1.setCellValue(c265.getText().toString().trim());
            rachq2.setCellValue(c266.getText().toString().trim());
            rachq3.setCellValue(c267.getText().toString().trim());

            HSSFRow rackqj = CONTAINER.getRow(128);
            HSSFCell racjq1 = rackqj.getCell(4);
            HSSFCell racjq2 = rackqj.getCell(5);
            HSSFCell racjq3 = rackqj.getCell(3);
            racjq1.setCellValue(c268.getText().toString().trim());
            racjq2.setCellValue(c269.getText().toString().trim());
            racjq3.setCellValue(c270.getText().toString().trim());

            HSSFRow rackwk = CONTAINER.getRow(134);
            HSSFCell rackw1 = rackwk.getCell(4);
            HSSFCell rackw2 = rackwk.getCell(5);
            HSSFCell rackw3 = rackwk.getCell(3);
            rackw1.setCellValue(c271.getText().toString().trim());
            rackw2.setCellValue(c272.getText().toString().trim());
            rackw3.setCellValue(c273.getText().toString().trim());

            HSSFRow rackel = CONTAINER.getRow(135);
            HSSFCell racle1 = rackel.getCell(4);
            HSSFCell racle2 = rackel.getCell(5);
            HSSFCell racle3 = rackel.getCell(3);
            racle1.setCellValue(c274.getText().toString().trim());
            racle2.setCellValue(c275.getText().toString().trim());
            racle3.setCellValue(c276.getText().toString().trim());

            HSSFRow rackdl = CONTAINER.getRow(136);
            HSSFCell racld1 = rackdl.getCell(4);
            HSSFCell racld2 = rackdl.getCell(5);
            HSSFCell racld3 = rackdl.getCell(3);
            racld1.setCellValue(c277.getText().toString().trim());
            racld2.setCellValue(c278.getText().toString().trim());
            racld3.setCellValue(c279.getText().toString().trim());

            HSSFRow rackcl = CONTAINER.getRow(137);
            HSSFCell raclc1 = rackcl.getCell(4);
            HSSFCell raclc2 = rackcl.getCell(5);
            HSSFCell raclc3 = rackcl.getCell(3);
            raclc1.setCellValue(c280.getText().toString().trim());
            raclc2.setCellValue(c281.getText().toString().trim());
            raclc3.setCellValue(c282.getText().toString().trim());

            HSSFRow rackxl = CONTAINER.getRow(138);
            HSSFCell raclx1 = rackxl.getCell(4);
            HSSFCell raclx2 = rackxl.getCell(5);
            HSSFCell raclx3 = rackxl.getCell(3);
            raclx1.setCellValue(c283.getText().toString().trim());
            raclx2.setCellValue(c284.getText().toString().trim());
            raclx3.setCellValue(c285.getText().toString().trim());

            HSSFRow rackml = CONTAINER.getRow(139);
            HSSFCell raclm1 = rackml.getCell(4);
            HSSFCell raclm2 = rackml.getCell(5);
            HSSFCell raclm3 = rackml.getCell(3);
            raclm1.setCellValue(c286.getText().toString().trim());
            raclm2.setCellValue(c287.getText().toString().trim());
            raclm3.setCellValue(c288.getText().toString().trim());

            HSSFRow racknl = CONTAINER.getRow(140);
            HSSFCell racln1 = racknl.getCell(4);
            HSSFCell racln2 = racknl.getCell(5);
            HSSFCell racln3 = racknl.getCell(3);
            racln1.setCellValue(c289.getText().toString().trim());
            racln2.setCellValue(c290.getText().toString().trim());
            racln3.setCellValue(c291.getText().toString().trim());

            HSSFRow rackfl = CONTAINER.getRow(141);
            HSSFCell raclf1 = rackfl.getCell(4);
            HSSFCell raclf2 = rackfl.getCell(5);
            HSSFCell raclf3 = rackfl.getCell(3);
            raclf1.setCellValue(c292.getText().toString().trim());
            raclf2.setCellValue(c293.getText().toString().trim());
            raclf3.setCellValue(c294.getText().toString().trim());



            myInputStream.close();
            FileOutputStream fos =new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename));
            workbook.write(fos);
            fos.close();
            Toast.makeText(context, "Changes are done.", Toast.LENGTH_SHORT).show();
        }
        catch (Exception e){e.printStackTrace();}
    }
}
