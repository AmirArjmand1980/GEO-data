package com.amirarjmand.pgd;

import android.content.Intent;
import android.os.Environment;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.util.SparseBooleanArray;
import android.view.View;
import android.widget.AdapterView;
import android.widget.ArrayAdapter;
import android.widget.Button;
import android.widget.ListView;
import android.widget.Toast;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.nio.channels.FileChannel;
import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class RR_Files extends AppCompatActivity {


    ListView lv;
    ArrayAdapter<String> adapter;
    List<String> arr;
    String[] fileName = {};
    String dirpath2,dirpath, listt;


    @Override
        protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(RR_Files.this);
        setContentView(R.layout.activity_rr__files);

        lv=findViewById(R.id.Flist);


        dirpath2 = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/RepairReport/Temp";
        dirpath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/RepairReport";

        arr = new ArrayList<String>(Arrays.asList(fileName));
        adapter = new ArrayAdapter<String>(this, R.layout.list_row, arr);
        lv.setAdapter(adapter);
        lv.setChoiceMode(ListView.CHOICE_MODE_MULTIPLE);
        lv.setItemsCanFocus(false);



        File[] files = new File(dirpath2).listFiles();

        for (int j=0; j<files.length; j++){

            listt=files[j].getName();
            arr.add(listt);
           }




        lv.setOnItemClickListener(new AdapterView.OnItemClickListener() {
            @Override
            public void onItemClick(AdapterView<?> parent, View view, int position, long id) {
                String fname=arr.get(position);
                Intent intent=new Intent(RR_Files.this, RepairReport.class);
                intent.putExtra("fname", fname);
                Toast.makeText(RR_Files.this, fname, Toast.LENGTH_SHORT).show();
                File source = new File(dirpath2, fname);
                File dest = new File(dirpath, "Repair Report.xls");

                    try {
                        copyFileUsingChannel(source, dest);


                    } catch (IOException e) {
                        e.printStackTrace();
                    }

           
                startActivity(intent);

            }
        });




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
    }

}
