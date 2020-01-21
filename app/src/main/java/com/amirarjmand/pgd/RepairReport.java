package com.amirarjmand.pgd;

import android.app.AlertDialog;
import android.content.Context;
import android.content.Intent;
import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Environment;
import android.support.annotation.NonNull;
import android.support.v4.content.FileProvider;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.view.WindowManager;
import android.widget.Button;
import android.widget.CalendarView;
import android.widget.EditText;
import android.widget.RadioButton;
import android.widget.RadioGroup;
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

public class RepairReport extends AppCompatActivity {

    EditText c1, c2, c3, c4, c5, c6, c7, c8, c9, c10, c11, c12, c13;
    private RadioGroup radioGroup, radioGroup2;
    private RadioButton radioButton1, radioButton2, radioButton3, radioButton4,radioButton5;
    Button save, fl,share, blank;
    CalendarView calendarView;
    String date, dirpath, dirpath2, radio1, radio2, filename;
    int  selectedId;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(RepairReport.this);
        getWindow().setSoftInputMode(WindowManager.LayoutParams.SOFT_INPUT_ADJUST_PAN);
        setContentView(R.layout.activity_repair_report);

        final Context context=RepairReport.this;
        save = findViewById(R.id.Save);
        fl=findViewById(R.id.FL);
        blank=findViewById(R.id.Blank);
        share=findViewById(R.id.Sharing);
        calendarView = findViewById(R.id.Calen);
        radioButton1=findViewById(R.id.Ok);
        radioButton2=findViewById(R.id.Waste);
        radioButton3=findViewById(R.id.Try);
        radioButton4=findViewById(R.id.Yes);
        radioButton5=findViewById(R.id.No);

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


        dirpath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/RepairReport";
        dirpath2 = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/RepairReport/Temp";

        calendarView.setOnDateChangeListener(new CalendarView.OnDateChangeListener() {
            @Override
            public void onSelectedDayChange(@NonNull CalendarView calendarView, int year, int month, int dayOfMonth) {
                date = dayOfMonth + "." + (month + 1) + "." + year;
                }
        });






        copyAsset("Repair Report.xls");
        ReadDataFromExcel(RepairReport.this, "Repair Report.xls");



        save.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                SaveDatainxcel(RepairReport.this, "Repair Report.xls");
                String device = c1.getText().toString();
                File source = new File(dirpath, "Repair Report.xls");
                String destinf = "Repair Report " + date + " (" + device;
                File dest = new File(dirpath2, destinf + ").xls");

                    try {
                        copyFileUsingChannel(source, dest);

                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                share.setOnClickListener(new View.OnClickListener() {
                    @Override
                    public void onClick(View v) {

                        ArrayList<Uri> uris = new ArrayList<Uri>();
                        String device = c1.getText().toString();
                        String destinf = "Repair Report " + date + " (" + device+").xls";

                        filename=destinf;
                        File filelocation = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/RepairReport/Temp", filename);
//                        Uri path1 = Uri.fromFile(filelocation);
                        Uri path1 = FileProvider.getUriForFile(context, context.getApplicationContext().getPackageName()+".provider",  filelocation );
                        uris.add(0,path1);
                        Intent emailIntent = new Intent(Intent.ACTION_SEND_MULTIPLE);
                        emailIntent .setType("vnd.android.cursor.dir/email");
                        String to[] = {"amir.arjmand@yahoo.com"};
                        emailIntent .putExtra(Intent.EXTRA_EMAIL, to);
                        emailIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
                        emailIntent .putExtra(Intent.EXTRA_SUBJECT, filename);
                        startActivity(Intent.createChooser(emailIntent , "Send file Via..."));

                    }
                });
                }

            });

        share.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                AlertDialog.Builder builder=new AlertDialog.Builder(RepairReport.this, R.style.MyAlertDialogStyle);
                builder.setTitle("Note");
                builder.setMessage("You have to save file before share it");
                builder.show();

            }
        });



      fl.setOnClickListener(new View.OnClickListener() {
          @Override
          public void onClick(View v) {
              startActivity(new Intent(RepairReport.this, RR_Files.class));
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
          }
      });




    }
    
    public void addListenerOnButton() {

        radioGroup = (RadioGroup) findViewById(R.id.radio);
        final int selectedId = radioGroup.getCheckedRadioButtonId();
        if (selectedId==-1){

            AlertDialog.Builder alert=new AlertDialog.Builder(RepairReport.this, R.style.MyAlertDialogStyle);
            alert.setTitle("Sorry");
            alert.setMessage("This device is repaired or not or needs to trying later?" +
                    "Please check the button");
            alert.setCancelable(true);
            alert.show();

        } else{
            radioButton1 = (RadioButton) findViewById(selectedId);
            radio1 = radioButton1.getText().toString();
        }}
    

    public void addListenerOnButton2() {

        radioGroup2 = (RadioGroup) findViewById(R.id.radio2);
        int selectedId = radioGroup2.getCheckedRadioButtonId();
        if (selectedId==-1){

            AlertDialog.Builder alert=new AlertDialog.Builder(RepairReport.this,R.style.MyAlertDialogStyle);
            alert.setTitle("Sorry");
            alert.setMessage("If this device is repaired, have you used spare part for it?" +
                    "Please check the button");
            
            alert.setCancelable(true);
            alert.show();

        } else{
            radioButton2 = (RadioButton) findViewById(selectedId);
            radio2 = radioButton2.getText().toString();
        }}





    private void ReadDataFromExcel(RepairReport Repair, String filename) {

        try {

            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/RepairReport/", filename);
            FileInputStream myInputStream = new FileInputStream(file);
            POIFSFileSystem poifsFileSystem = new POIFSFileSystem(myInputStream);

            HSSFWorkbook inv = new HSSFWorkbook(poifsFileSystem);
            HSSFSheet BUS_SYSTEM = inv.getSheetAt(0);

            HSSFRow Condenser = BUS_SYSTEM.getRow(5);
            HSSFCell ConCell = Condenser.getCell(2);
            c1.setText(ConCell.toString());

            HSSFRow ups = BUS_SYSTEM.getRow(5);
            HSSFCell ConCel = ups.getCell(8);
            c2.setText(ConCel.toString());

            HSSFRow upss = BUS_SYSTEM.getRow(7);
            HSSFCell Con = upss.getCell(2);
            c3.setText(Con.toString());

            HSSFRow up = BUS_SYSTEM.getRow(7);
            HSSFCell Co = up.getCell(8);
            c4.setText(Co.toString());

            HSSFRow upse = BUS_SYSTEM.getRow(9);
            HSSFCell ConCele = upse.getCell(2);
            c5.setText(ConCele.toString());

            HSSFRow upis = BUS_SYSTEM.getRow(12);
            HSSFCell ConiCel = upis.getCell(0);
            c6.setText(ConiCel.toString());

            HSSFRow upiis = BUS_SYSTEM.getRow(13);
            HSSFCell ConiiCel = upiis.getCell(0);
            c7.setText(ConiiCel.toString());

            HSSFRow uups = BUS_SYSTEM.getRow(20);
            HSSFCell ConuCel = uups.getCell(0);
            c8.setText(ConuCel.toString());

            HSSFRow uuups = BUS_SYSTEM.getRow(21);
            HSSFCell ConuuCel = uuups.getCell(0);
            c9.setText(ConuuCel.toString());

            HSSFRow upts = BUS_SYSTEM.getRow(22);
            HSSFCell ContCel = upts.getCell(0);
            c10.setText(ContCel.toString());

            HSSFRow upsx = BUS_SYSTEM.getRow(23);
            HSSFCell ConCelx = upsx.getCell(0);
            c11.setText(ConCelx.toString());

            HSSFRow upws = BUS_SYSTEM.getRow(24);
            HSSFCell ConwCel = upws.getCell(0);
            c12.setText(ConwCel.toString());

            HSSFRow uupws = BUS_SYSTEM.getRow(42);
            HSSFCell ConuwCel = uupws.getCell(3);
            c13.setText(ConuwCel.toString());



        } catch (Exception e) {
            e.printStackTrace();
        }


    }

    private void SaveDatainxcel(Context context, String filename) {

        try {
            File file = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/RepairReport/", filename);
            FileInputStream myInputStream = new FileInputStream(file);

            HSSFWorkbook workbook = new HSSFWorkbook(myInputStream);
            HSSFSheet BUS_SYSTEM = workbook.getSheetAt(0);

            HSSFRow cond = BUS_SYSTEM.getRow(5);
            HSSFCell con1 = cond.getCell(2);
            con1.setCellValue(c1.getText().toString().trim());

            HSSFRow cond1 = BUS_SYSTEM.getRow(5);
            HSSFCell con11 = cond1.getCell(8);
            con11.setCellValue(c2.getText().toString().trim());

            HSSFRow cond11 = BUS_SYSTEM.getRow(7);
            HSSFCell con111 = cond11.getCell(2);
            con111.setCellValue(c3.getText().toString().trim());

            HSSFRow cond111 = BUS_SYSTEM.getRow(7);
            HSSFCell con1111 = cond111.getCell(8);
            con1111.setCellValue(c4.getText().toString().trim());

            HSSFRow cond2 = BUS_SYSTEM.getRow(9);
            HSSFCell con2 = cond2.getCell(2);
            con2.setCellValue(c5.getText().toString().trim());

            HSSFRow cond22 = BUS_SYSTEM.getRow(12);
            HSSFCell con22 = cond22.getCell(0);
            con22.setCellValue(c6.getText().toString().trim());

            HSSFRow cond222 = BUS_SYSTEM.getRow(13);
            HSSFCell con222 = cond222.getCell(0);
            con222.setCellValue(c7.getText().toString().trim());

            HSSFRow cond3 = BUS_SYSTEM.getRow(20);
            HSSFCell con3 = cond3.getCell(0);
            con3.setCellValue(c8.getText().toString().trim());

            HSSFRow cond33 = BUS_SYSTEM.getRow(21);
            HSSFCell con33 = cond33.getCell(0);
            con33.setCellValue(c9.getText().toString().trim());

            HSSFRow cond333 = BUS_SYSTEM.getRow(22);
            HSSFCell con333 = cond333.getCell(0);
            con333.setCellValue(c10.getText().toString().trim());

            HSSFRow cond4 = BUS_SYSTEM.getRow(23);
            HSSFCell con4 = cond4.getCell(0);
            con4.setCellValue(c11.getText().toString().trim());

            HSSFRow cond44 = BUS_SYSTEM.getRow(24);
            HSSFCell con44 = cond44.getCell(0);
            con44.setCellValue(c12.getText().toString().trim());

            addListenerOnButton();
            HSSFRow cond55 = BUS_SYSTEM.getRow(34);
            HSSFCell con55 = cond55.getCell(1);
            con55.setCellValue(radio1);


                addListenerOnButton2();
                HSSFRow cond5 = BUS_SYSTEM.getRow(40);
                HSSFCell con5 = cond5.getCell(3);

                con5.setCellValue(radio2);


                HSSFRow cond444 = BUS_SYSTEM.getRow(42);
                HSSFCell con444 = cond444.getCell(3);
                con444.setCellValue(c13.getText().toString().trim());

                HSSFRow cond4444 = BUS_SYSTEM.getRow(45);
                HSSFCell con4444 = cond4444.getCell(3);
                con4444.setCellValue(date);




            myInputStream.close();
            FileOutputStream fos = new FileOutputStream(new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/RepairReport/", filename));
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

        File RR = new File(dirpath, "Repair Report.xls");

        if (!RR.exists()) {

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
        }
    }

    private boolean listAssetFiles(String path) {

        String[] list;
        try {
            list = getAssets().list(path);
            if (list.length > 0) {
                // This is a folder
                for (String file : list) {
                    if (!listAssetFiles(path + "/" + file)) return false;
                    else {
                        // This is a file
                        // TODO: add file name to an array list
                    }
                }
            }
        } catch (IOException e) {
            return false;
        }

        return true;
    }


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
        }
    }}








