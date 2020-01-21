package com.amirarjmand.pgd;

import android.content.Intent;
import android.content.res.AssetManager;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.ListView;
import android.widget.Toast;

import java.io.File;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.InputStream;
import java.io.OutputStream;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class Unit_List extends AppCompatActivity {
    ListView lv;
    ArrayAdapter<String> adapter;
    List<String> arr;
    String[] language = {"Unit 101", "Unit 122", "Unit 123", "Unit 124", "Unit 125", "Unit 127", "Unit 128", "Unit 129", "Unit 130", "Unit 131", "Unit 132", "Unit 133", "Unit 134", "Unit 135", "Unit 136"};
    String lvStr, UnitNo;
    String dirpath, filename;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(Unit_List.this);
        setContentView(R.layout.activity_unit__list);


        lv = (ListView) findViewById(R.id.listView1);
        arr = new ArrayList<String>(Arrays.asList(language));
        adapter = new ArrayAdapter<String>(this, R.layout.list_row, arr);
        lv.setAdapter(adapter);

       lv.setOnItemClickListener(new AdapterView.OnItemClickListener() {
           @Override
           public void onItemClick(AdapterView<?> parent, View view, int position, long id) {

               lvStr = arr.get(position);
               UnitNo = lvStr.replace("Unit ", "").trim();
               Intent intent = new Intent(Unit_List.this, Content.class);
               intent.putExtra("uno", UnitNo);
               copyAsset("Inventory List unit "+UnitNo+".xls");
               copyAsset("Unit "+UnitNo+" Shortage.xls");
               startActivity(intent);
           }
       });




    }

    private void copyAsset(String filename) {
         dirpath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/Inventory";
        File dir = new File(dirpath);
        if (!dir.exists()) {
            dir.mkdirs();
        }

        File invent = new File(dirpath, "Inventory List unit " + UnitNo + ".xls");

        if (!invent.exists()) {

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

        File shortage=new File(dirpath, "Unit "+UnitNo+" Shortage.xls");
       if (!shortage.exists()){

            AssetManager assetManager1 = getAssets();
            InputStream in1 = null;
            OutputStream out1 = null;
            try {
                in1 = assetManager1.open(filename);
                File outfile1 = new File(dirpath, filename);
                out1 = new FileOutputStream(outfile1);
                copyfile(in1, out1);


            } catch (IOException e) {
                e.printStackTrace();

            } finally {
                if (in1 != null) {
                    try {
                        in1.close();
                    } catch (IOException e) {
                        e.printStackTrace();
                    }
                }
                if (out1 != null) {
                    try {
                        out1.close();

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
    
}


