package com.amirarjmand.pgd;

import android.app.AlertDialog;
import android.app.DatePickerDialog;
import android.app.Dialog;
import android.content.Context;
import android.content.Intent;
import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Bundle;
import android.os.Environment;
import android.support.v4.content.FileProvider;
import android.support.v7.app.AppCompatActivity;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.DatePicker;
import android.widget.EditText;
import android.widget.ImageView;
import android.widget.RadioButton;
import android.widget.RadioGroup;
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
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.Calendar;




public class DER extends AppCompatActivity {

    TextView datedisplay, today;
    DatePicker date_picker;
    ImageView cal;
    EditText c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13, c14, c15, c16, c17, c18, c19, c20, c21, c22;
    RadioGroup radioGroup1, radioGroup2, radioGroup3;
    Button save, fl, share, blank;
    int year,month, day,radiobuttonId1,radiobuttonId2,radiobuttonId3,idx1,idx2,idx3;
    String date, dirpath, dirpath2,filename;
    Context context;
    static final int DATE_DIALOG_ID = 100;



    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(DER.this);
        setContentView(R.layout.activity_der);



        datedisplay = findViewById(R.id.DateDisplay);
        cal = findViewById(R.id.Cal);
        today = findViewById(R.id.Today);
        radioGroup1=findViewById(R.id.RG1);
        radioGroup2=findViewById(R.id.RG2);
        radioGroup3=findViewById(R.id.RG3);

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
        save = findViewById(R.id.Save);
        fl = findViewById(R.id.FL);
        blank = findViewById(R.id.Blank);
        share = findViewById(R.id.Sharing);

        context=DER.this;

        dirpath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/DER";
        dirpath2 = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/DER/Temp";


        copyAsset("DER.xls");
        ReadDataFromExcel(DER.this, "DER.xls");



        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                date = datedisplay.getText().toString();
                SaveDatainxcel(DER.this, "DER.xls");
                final String unitno = c2.getText().toString();
                File source = new File(dirpath, "DER.xls");
                String destinf = "DER " + unitno + " (" + date;
                File dest = new File(dirpath2, destinf + ").xls");
                try {copyFileUsingChannel(source, dest);} catch (IOException e) {e.printStackTrace(); }
                ClearBoxes(DER.this, "DER.xls");




                share.setOnClickListener(new View.OnClickListener() {
                    @Override
                    public void onClick(View v) {

                        ArrayList<Uri> uris = new ArrayList<Uri>();
                        String device = c1.getText().toString();
                        String destinf = "DER " + unitno + " (" + date + ").xls";
                        filename = destinf;
                        File filelocation = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/DER/Temp", filename);
                        Uri path1 = FileProvider.getUriForFile(context, context.getApplicationContext().getPackageName() + ".provider", filelocation);
                        uris.add(0, path1);
                        Intent emailIntent = new Intent(Intent.ACTION_SEND_MULTIPLE);
                        emailIntent.setType("vnd.android.cursor.dir/email");
                        String to[] = {"amir.arjmand@yahoo.com"};
                        emailIntent.putExtra(Intent.EXTRA_EMAIL, to);
                        emailIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
                        emailIntent.putExtra(Intent.EXTRA_SUBJECT, filename);
                        startActivity(Intent.createChooser(emailIntent, "Send file Via..."));

                    }
                });
            }

        });

        share.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                AlertDialog.Builder builder = new AlertDialog.Builder(DER.this, R.style.MyAlertDialogStyle);
                builder.setTitle("Note");
                builder.setMessage("You have to save file before share it");
                builder.show();

            }
        });


        fl.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                startActivity(new Intent(DER.this, DER_Files.class));
            }
        });


        blank.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                c1.setText("");
                c2.setText("");
                c3.setText("");
                c4.setText("");
                c5.setText("");
                c6.setText("");
                c7.setText("");
                c8.setText("");
                c9.setText("");
                c10.setText("");
                c11.setText("");
                c12.setText("");
                c13.setText("");
                c14.setText("");
                c15.setText("");
                c16.setText("");
                c17.setText("");
                c18.setText("");
                c19.setText("");
                c20.setText("");
                c21.setText("");
                c22.setText("");

            }
        });


        today.setOnClickListener(new OnClickListener() {
            @Override
            public void onClick(View v) {
                setCurrentDate();

            }
        });

        cal.setOnClickListener(new OnClickListener() {
            @Override
            public void onClick(View v) {
                addButtonListener();

            }
        });


    }


    public void setCurrentDate() {
        date_picker = (DatePicker) findViewById(R.id.date_picker);
        final Calendar calendar = Calendar.getInstance();
        year = calendar.get(Calendar.YEAR);
        month = calendar.get(Calendar.MONTH);
        day = calendar.get(Calendar.DAY_OF_MONTH);

        datedisplay.setText(new StringBuilder().append(day).append(".").append(month + 1).append(".").append(year));
        date_picker.init(year, month, day, null);
    }





    public void addButtonListener() {
        final Calendar calendar1 = Calendar.getInstance();

        year = calendar1.get(Calendar.YEAR);
        month = calendar1.get(Calendar.MONTH);
        day = calendar1.get(Calendar.DAY_OF_MONTH);
        datedisplay.setText(new StringBuilder().append(day).append(".").append(month + 1).append(".").append(year));
        date_picker.init(year, month, day, null);
        showDialog(DATE_DIALOG_ID);}



    @Override
    protected Dialog onCreateDialog(int id) {
        switch (id) {
            case DATE_DIALOG_ID:
                DatePickerDialog datePickerDialog = new DatePickerDialog(this, datePickerListener, year, month, day);
                return datePickerDialog;
        }
        return null;}





    private DatePickerDialog.OnDateSetListener datePickerListener = new DatePickerDialog.OnDateSetListener() {

        // when dialog box is closed, below method will be called.
        public void onDateSet(DatePicker view, int selectedYear, int selectedMonth, int selectedDay) {
            year = selectedYear;
            month = selectedMonth;
            day = selectedDay;
            datedisplay.setText(new StringBuilder().append(day).append(".").append(month + 1).append(".").append(year));
            date_picker.init(year, month, day, null); }
    };





    private void ReadDataFromExcel(DER DER, String filename) {

        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/DER/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook inv = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet BUS_SYSTEM = inv.getSheetAt(0);

            HSSFRow Condenser = BUS_SYSTEM.getRow(4);
            HSSFCell ConCell = Condenser.getCell(3);
            c1.setText(ConCell.toString());

            HSSFRow ups = BUS_SYSTEM.getRow(4);
            HSSFCell ConCel = ups.getCell(7);
            c2.setText(ConCel.toString());

            HSSFRow upss = BUS_SYSTEM.getRow(8);
            HSSFCell Con = upss.getCell(1);
            c3.setText(Con.toString());

            HSSFRow up = BUS_SYSTEM.getRow(8);
            HSSFCell Co = up.getCell(10);
            c4.setText(Co.toString());

            HSSFRow upse = BUS_SYSTEM.getRow(12);
            HSSFCell ConCele = upse.getCell(7);
            c5.setText(ConCele.toString());

            HSSFRow upis = BUS_SYSTEM.getRow(14);
            HSSFCell ConiCel = upis.getCell(13);
            c6.setText(ConiCel.toString());

            HSSFRow upiis = BUS_SYSTEM.getRow(18);
            HSSFCell ConiiCel = upiis.getCell(3);
            c7.setText(ConiiCel.toString());

            HSSFRow uups = BUS_SYSTEM.getRow(18);
            HSSFCell ConuCel = uups.getCell(13);
            c8.setText(ConuCel.toString());

            HSSFRow uuups = BUS_SYSTEM.getRow(23);
            HSSFCell ConuuCel = uuups.getCell(1);
            c9.setText(ConuuCel.toString());

            HSSFRow upts = BUS_SYSTEM.getRow(24);
            HSSFCell ContCel = upts.getCell(1);
            c10.setText(ContCel.toString());

            HSSFRow upsx = BUS_SYSTEM.getRow(25);
            HSSFCell ConCelx = upsx.getCell(1);
            c11.setText(ConCelx.toString());

            HSSFRow upws = BUS_SYSTEM.getRow(26);
            HSSFCell ConwCel = upws.getCell(1);
            c12.setText(ConwCel.toString());

            HSSFRow uupws = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCel = uupws.getCell(1);
            c13.setText(ConuwCel.toString());

            HSSFRow uufpws = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCelf = uufpws.getCell(1);
            c14.setText(ConuwCelf.toString());
            HSSFRow uupwsf = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCefl = uupwsf.getCell(1);
            c15.setText(ConuwCefl.toString());

            HSSFRow uupwsv = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCelv = uupwsv.getCell(1);
            c16.setText(ConuwCelv.toString());

            HSSFRow uupwsc = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCelc = uupwsc.getCell(1);
            c17.setText(ConuwCelc.toString());

            HSSFRow uupwsz = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCelz = uupwsz.getCell(1);
            c18.setText(ConuwCelz.toString());

            HSSFRow uupwse = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCele = uupwse.getCell(1);
            c19.setText(ConuwCele.toString());

            HSSFRow uupwsr = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCelr = uupwsr.getCell(1);
            c20.setText(ConuwCelr.toString());

            HSSFRow uupwst = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCelt = uupwst.getCell(1);
            c21.setText(ConuwCelt.toString());

            HSSFRow uupwsy = BUS_SYSTEM.getRow(27);
            HSSFCell ConuwCely = uupwsy.getCell(1);
            c22.setText(ConuwCely.toString());


        } catch (Exception e) {
            e.printStackTrace();
        }


    }




    private void SaveDatainxcel(Context context, String filename) {

        try {
            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/DER/", filename);
            FileInputStream myInputStream = new FileInputStream(file);

            HSSFWorkbook workbook = new HSSFWorkbook(myInputStream);
            HSSFSheet BUS_SYSTEM = workbook.getSheetAt(0);

            HSSFRow cond = BUS_SYSTEM.getRow(4);
            HSSFCell con1 = cond.getCell(3);
            con1.setCellValue(c1.getText().toString().trim());

            HSSFRow cond1 = BUS_SYSTEM.getRow(4);
            HSSFCell con11 = cond1.getCell(7);
            con11.setCellValue(c2.getText().toString().trim());

            HSSFRow cond11 = BUS_SYSTEM.getRow(8);
            HSSFCell con111 = cond11.getCell(1);
            con111.setCellValue(c3.getText().toString().trim());

            HSSFRow cond111 = BUS_SYSTEM.getRow(8);
            HSSFCell con1111 = cond111.getCell(10);
            con1111.setCellValue(c4.getText().toString().trim());

            HSSFRow cond2 = BUS_SYSTEM.getRow(12);
            HSSFCell con2 = cond2.getCell(7);
            con2.setCellValue(c5.getText().toString().trim());

            HSSFRow cond22 = BUS_SYSTEM.getRow(14);
            HSSFCell con22 = cond22.getCell(13);
            con22.setCellValue(c6.getText().toString().trim());

            HSSFRow cond222 = BUS_SYSTEM.getRow(18);
            HSSFCell con222 = cond222.getCell(3);
            con222.setCellValue(c7.getText().toString().trim());

            HSSFRow cond3 = BUS_SYSTEM.getRow(18);
            HSSFCell con3 = cond3.getCell(13);
            con3.setCellValue(c8.getText().toString().trim());

            HSSFRow cond33 = BUS_SYSTEM.getRow(23);
            HSSFCell con33 = cond33.getCell(1);
            con33.setCellValue(c9.getText().toString().trim());

            HSSFRow cond333 = BUS_SYSTEM.getRow(24);
            HSSFCell con333 = cond333.getCell(1);
            con333.setCellValue(c10.getText().toString().trim());

            HSSFRow cond4 = BUS_SYSTEM.getRow(25);
            HSSFCell con4 = cond4.getCell(1);
            con4.setCellValue(c11.getText().toString().trim());

            HSSFRow cond44 = BUS_SYSTEM.getRow(26);
            HSSFCell con44 = cond44.getCell(1);
            con44.setCellValue(c12.getText().toString().trim());

            HSSFRow cond444 = BUS_SYSTEM.getRow(27);
            HSSFCell con444 = cond444.getCell(1);
            con444.setCellValue(c13.getText().toString().trim());

            HSSFRow cond1444 = BUS_SYSTEM.getRow(27);
            HSSFCell con1444 = cond1444.getCell(1);
            con1444.setCellValue(c14.getText().toString().trim());

            HSSFRow qcond444 = BUS_SYSTEM.getRow(27);
            HSSFCell qcon444 = qcond444.getCell(1);
            qcon444.setCellValue(c15.getText().toString().trim());

            HSSFRow zcond444 = BUS_SYSTEM.getRow(27);
            HSSFCell zcon444 = zcond444.getCell(1);
            zcon444.setCellValue(c16.getText().toString().trim());

            HSSFRow xcond444 = BUS_SYSTEM.getRow(27);
            HSSFCell xcon444 = xcond444.getCell(1);
            xcon444.setCellValue(c17.getText().toString().trim());

            HSSFRow bcond444 = BUS_SYSTEM.getRow(27);
            HSSFCell bcon444 = bcond444.getCell(1);
            bcon444.setCellValue(c18.getText().toString().trim());

            HSSFRow tcond444 = BUS_SYSTEM.getRow(27);
            HSSFCell tcon444 = tcond444.getCell(1);
            tcon444.setCellValue(c19.getText().toString().trim());

            HSSFRow dcond444 = BUS_SYSTEM.getRow(27);
            HSSFCell dcon444 = dcond444.getCell(1);
            dcon444.setCellValue(c20.getText().toString().trim());

            HSSFRow scond444 = BUS_SYSTEM.getRow(27);
            HSSFCell scon444 = scond444.getCell(1);
            scon444.setCellValue(c21.getText().toString().trim());

            HSSFRow icond444 = BUS_SYSTEM.getRow(27);
            HSSFCell icon444 = icond444.getCell(1);
            icon444.setCellValue(c22.getText().toString().trim());

            HSSFRow eicond444 = BUS_SYSTEM.getRow(41);
            HSSFCell eicon444 = eicond444.getCell(3);
            eicon444.setCellValue(c3.getText().toString().trim());

            HSSFRow zeicond444 = BUS_SYSTEM.getRow(41);
            HSSFCell zeicon444 = zeicond444.getCell(10);
            zeicon444.setCellValue(c4.getText().toString().trim());

            HSSFRow ond444 = BUS_SYSTEM.getRow(41);
            HSSFCell on444 = ond444.getCell(16);
            on444.setCellValue("Alireza Mardan");

            Radio1();
            if (idx1 == 0){
                HSSFRow ond1444 = BUS_SYSTEM.getRow(38);
                HSSFCell on1444 = ond1444.getCell(4);
                on1444.setCellValue("X");
            } else {
                HSSFRow ond21444 = BUS_SYSTEM.getRow(38);
                HSSFCell on21444 = ond21444.getCell(6);
                on21444.setCellValue("X");
            }

            Radio2();
            if (idx2 == 0){
                HSSFRow o4 = BUS_SYSTEM.getRow(38);
                HSSFCell on4 = o4.getCell(13);
                on4.setCellValue("X");
            } else {
                HSSFRow on4z = BUS_SYSTEM.getRow(38);
                HSSFCell o4z = on4z.getCell(15);
                o4z.setCellValue("X");
            }

            Radio3();
            if (idx3 == 0){
                HSSFRow ss = BUS_SYSTEM.getRow(39);
                HSSFCell sss = ss.getCell(13);
                sss.setCellValue("X");
            } else {
                HSSFRow dd = BUS_SYSTEM.getRow(39);
                HSSFCell ddd = dd.getCell(15);
                ddd.setCellValue("X");
            }

            HSSFRow cond4444 = BUS_SYSTEM.getRow(4);
            HSSFCell con4444 = cond4444.getCell(11);
            con4444.setCellValue(date);

            myInputStream.close();
            FileOutputStream fos = new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/DER/", filename));
            workbook.write(fos);
            fos.close();
            Toast.makeText(context, "Changes are done.", Toast.LENGTH_SHORT).show();
        } catch (Exception e) {
            e.printStackTrace();
        }
    }





    private void copyAsset(String filename) {
        File dir = new File(dirpath);
        File dir2 = new File(dirpath2);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        if (!dir2.exists()) {
            dir2.mkdir();
        }

        File der = new File(dirpath, "DER.xls");

        if (!der.exists()) {

            AssetManager assetManager = getAssets();
            InputStream in = null;
            OutputStream out = null;
            try {
                in = assetManager.open(filename);
                File outfile = new File(dirpath, filename);
                out = new FileOutputStream(outfile);
                copyfile(in, out);
                Toast.makeText(this, "File saved in storage", Toast.LENGTH_SHORT).show();

            } catch (IOException e) {
                e.printStackTrace();
                Toast.makeText(this, "Saving Failed", Toast.LENGTH_SHORT).show();
            } finally {
                if (in != null) {
                    try {
                        in.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if (out != null) {
                    try {
                        out.close();

                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
            }
        }
    }





    private void copyfile(InputStream in, OutputStream out) throws IOException {
        byte[] buffer = new byte[1024];
        int read;
        while ((read = in.read(buffer)) != -1) {
            out.write(buffer, 0, read);
        } }





    private boolean listAssetFiles(String path) {

        String[] list;
        try {
            list = getAssets().list(path);
            if (list.length > 0) {
                for (String file : list) {
                    if (!listAssetFiles(path + "/" + file)) return false;
                    else {}}}
        } catch (IOException e) {
            return false;}
        return true; }






    private static void copyFileUsingChannel(File source, File dest) throws IOException {
        FileChannel sourceChannel = null;
        FileChannel destChannel = null;
        try {
            sourceChannel = new FileInputStream(source).getChannel();
            destChannel = new FileOutputStream(dest).getChannel();
            destChannel.transferFrom(sourceChannel, 0, sourceChannel.size());
        } finally {
            sourceChannel.close();
            destChannel.close();
        }}


        private void Radio1 (){
        radiobuttonId1=radioGroup1.getCheckedRadioButtonId();
        View radioButton = radioGroup1.findViewById(radiobuttonId1);
        idx1=radioGroup1.indexOfChild(radioButton) ;
        }

         private void Radio2 (){
        radiobuttonId2=radioGroup2.getCheckedRadioButtonId();
        View radioButton2 = radioGroup2.findViewById(radiobuttonId2);
        idx2=radioGroup2.indexOfChild(radioButton2) ;
        }

          private void Radio3 (){
        radiobuttonId3=radioGroup3.getCheckedRadioButtonId();
        View radioButton3 = radioGroup3.findViewById(radiobuttonId3);
        idx3=radioGroup3.indexOfChild(radioButton3) ;
         }

         private void ClearBoxes(Context context, String filename){

             try {
                 File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/DER/", filename);
                 FileInputStream myInputStream = new FileInputStream(file);

                 HSSFWorkbook workbook = new HSSFWorkbook(myInputStream);
                 HSSFSheet BUS_SYSTEM = workbook.getSheetAt(0);

                 HSSFRow ond1444 = BUS_SYSTEM.getRow(38);
                 HSSFCell on1444 = ond1444.getCell(4);
                 on1444.setCellValue("");

                 HSSFRow ond21444 = BUS_SYSTEM.getRow(38);
                 HSSFCell on21444 = ond21444.getCell(6);
                 on21444.setCellValue("");

                 HSSFRow o4 = BUS_SYSTEM.getRow(38);
                 HSSFCell on4 = o4.getCell(13);
                 on4.setCellValue("");

                 HSSFRow on4z = BUS_SYSTEM.getRow(38);
                 HSSFCell o4z = on4z.getCell(15);
                 o4z.setCellValue("");

                 HSSFRow ss = BUS_SYSTEM.getRow(39);
                 HSSFCell sss = ss.getCell(13);
                 sss.setCellValue("");

                 HSSFRow dd = BUS_SYSTEM.getRow(39);
                 HSSFCell ddd = dd.getCell(15);
                 ddd.setCellValue("");


                 myInputStream.close();
                 FileOutputStream fos = new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/DER/", filename));
                 workbook.write(fos);
                 fos.close();
             } catch (Exception e) {
                 e.printStackTrace();
             }

         }


}



