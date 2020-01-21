package com.amirarjmand.pgd;

import android.content.Context;
import android.content.Intent;
import android.content.res.AssetManager;
import android.net.Uri;
import android.os.Environment;
import android.support.v4.content.FileProvider;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.ImageView;
import android.widget.ListView;
import android.widget.Toast;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.nio.channels.FileChannel;
import java.text.SimpleDateFormat;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.Date;
import java.util.List;
import java.util.Locale;

public class Unit_Loc_List extends AppCompatActivity {

    ListView lv;
    ArrayAdapter<String> adapter;
    List<String> arr;
    String[] language = {"Unit 101", "Unit 122", "Unit 123", "Unit 124", "Unit 125", "Unit 127", "Unit 128", "Unit 129", "Unit 130", "Unit 131", "Unit 132", "Unit 133", "Unit 134", "Unit 135", "Unit 136"};
    String lvStr, UnitNo;
    String dirpath,dirpath2, filename, datea ;
    ImageView sharing;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(Unit_Loc_List.this);
        setContentView(R.layout.activity_unit__loc__list);

        final Context context=Unit_Loc_List.this;

        sharing=findViewById(R.id.Sharing);

        dirpath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/Unit_Location";
        dirpath2 = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/Unit_Location/Temp";



        SimpleDateFormat sdf = new SimpleDateFormat("dd_MM_yyyy", Locale.getDefault());

        final String currentDateandTime = sdf.format(new Date());
        datea=sdf.format(new Date());

        lv = (ListView) findViewById(R.id.listView1);
        arr = new ArrayList<String>(Arrays.asList(language));
        adapter = new ArrayAdapter<String>(this, R.layout.list_row, arr);
        lv.setAdapter(adapter);

        lv.setOnItemClickListener(new AdapterView.OnItemClickListener() {
            @Override
            public void onItemClick(AdapterView<?> parent, View view, int position, long id) {

                lvStr = arr.get(position);
                UnitNo = lvStr.replace("Unit ", "").trim();
                Intent intent = new Intent(Unit_Loc_List.this, ULOC_Form.class);
                intent.putExtra("uno", UnitNo);
                copyAsset("Units Location.xls");
                startActivity(intent);
                finish();
            }
        });

        sharing.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                File source=new File(dirpath, "Units Location.xls");
                File dest=new File(dirpath2, "Units Location "+currentDateandTime+".xls");
                try {
                    copyFileUsingChannel(source, dest);

                } catch (IOException e) {
                    e.printStackTrace();
                }

                ArrayList<Uri> uris = new ArrayList<Uri>();
                filename="Units Location "+currentDateandTime+".xls";
                File filelocation = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Unit_Location/Temp", filename);
//                Uri path1 = Uri.fromFile(filelocation);
                Uri path1 = FileProvider.getUriForFile(context, context.getApplicationContext().getPackageName()+".provider",  filelocation );
                uris.add(0,path1);
                Intent emailIntent = new Intent(Intent.ACTION_SEND_MULTIPLE);
                emailIntent .setType("vnd.android.cursor.dir/email");
                String to[] = {"amir.arjmand@yahoo.com"};
                emailIntent .putExtra(Intent.EXTRA_EMAIL, to);
                emailIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
                emailIntent .putExtra(Intent.EXTRA_SUBJECT, "Units Location "+currentDateandTime);
                startActivity(Intent.createChooser(emailIntent , "Send file Via..."));

            }
        });






    }

    private void copyAsset(String filename) {


        File dir = new File(dirpath);
        File dir2=new File(dirpath2);
        if (!dir.exists()) {
            dir.mkdirs();
        }
        if (!dir2.exists()){
            dir2.mkdir();
        }

        File uloc = new File(dirpath, "Units Location.xls");

        if (!uloc.exists()) {

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
        }finally{
            sourceChannel.close();
            destChannel.close();
        }
    }
}


