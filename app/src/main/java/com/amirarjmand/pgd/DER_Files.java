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

public class DER_Files extends AppCompatActivity {


    ListView lv;
    ArrayAdapter<String> adapter;
    List<String> arr;
    String[] fileName = {};
    String dirpath2,dirpath, listt;
    Button sharing;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(DER_Files.this);
        setContentView(R.layout.activity_der__files);

        lv=findViewById(R.id.Flist);
        sharing=findViewById(R.id.Share);

        dirpath2 = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/DER/Temp";
        dirpath = Environment.getExternalStorageDirectory().getAbsolutePath() + "/GEO-data/DER";

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
                Intent intent=new Intent(DER_Files.this, DER.class);
                intent.putExtra("fname", fname);
                Toast.makeText(DER_Files.this, fname, Toast.LENGTH_SHORT).show();
                File source = new File(dirpath2, fname);
                File dest = new File(dirpath, "DER.xls");

                try {
                    copyFileUsingChannel(source, dest);


                } catch (IOException e) {
                    e.printStackTrace();
                }


                startActivity(intent);

            }
        });


        sharing.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                SparseBooleanArray checked = lv.getCheckedItemPositions();
                ArrayList<String> selectedItems = new ArrayList<String>();
                for (int i = 0; i < checked.size(); i++) {
                    // Item position in adapter
                    int position = checked.keyAt(i);
                    // Add sport if it is checked i.e.) == TRUE!
                    if (checked.valueAt(i))
                        selectedItems.add(adapter.getItem(position));
                }
                String[] outputStrArr = new String[selectedItems.size()];

                for (int i = 0; i < selectedItems.size(); i++) {
                    outputStrArr[i] = selectedItems.get(i);
                }
                Intent intent = new Intent(getApplicationContext(),
                        ResultActivity.class);

                // Create a bundle object
                Bundle b = new Bundle();
                b.putStringArray("selectedItems", outputStrArr);

                // Add the bundle to the intent.
                intent.putExtras(b);

                // start the ResultActivity
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
