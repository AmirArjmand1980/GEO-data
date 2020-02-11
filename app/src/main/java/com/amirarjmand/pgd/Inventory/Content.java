package com.amirarjmand.pgd.Inventory;

import android.content.Context;
import android.content.Intent;
import android.net.Uri;
import android.support.v4.content.FileProvider;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.ImageView;
import android.widget.TextView;

import com.amirarjmand.pgd.R;
import com.amirarjmand.pgd.Unit_List;
import com.amirarjmand.pgd.noTitle;

import java.io.File;
import java.util.ArrayList;

public class Content extends AppCompatActivity {

    TextView contents, container, bus, gas, alarm, safety, sensor, cable, geolog, riginf, test, tools, stationary, accesory, description, shortage;
    String UnitNo;
    String filename1, filename2;

    ImageView sharing;

    private static final int MY_PERMISSION_REQUEST_STORAGE = 1;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(Content.this);
        setContentView(R.layout.activity_content);

        geolog = findViewById(R.id.Geo);
        contents = findViewById(R.id.textView);
        container = findViewById(R.id.container);
        bus = findViewById(R.id.BUS);
        gas = findViewById(R.id.GAS);
        alarm = findViewById(R.id.Alarm);
        sensor = findViewById(R.id.Sensor);
        cable = findViewById(R.id.Cable);
        riginf = findViewById(R.id.RIG);
        test = findViewById(R.id.Test);
        safety = findViewById(R.id.Safety);
        tools = findViewById(R.id.Tool);
        stationary = findViewById(R.id.Stationary);
        accesory = findViewById(R.id.Other);
        sharing=findViewById(R.id.Sharing);
        description=findViewById(R.id.Descript);
        shortage=findViewById(R.id.Short);

        final Context context=Content.this;

        filename1 = "Inventory List unit " + UnitNo + ".xls";
        filename2= "Unit "+UnitNo+" Shortage.xls";
        Bundle Uno = getIntent().getExtras();
        UnitNo = Uno.getString("uno");
        contents.setText("Select a content of unit " + UnitNo);

//        SaveDatainxcel(Content.this, "Inventory List unit "+UnitNo+".xls");

        container.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intent = new Intent(Content.this, Container.class);
                intent.putExtra("uno", UnitNo);
                startActivity(intent);
            }
        });


        bus.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentBus = new Intent(Content.this, BUS.class);
                intentBus.putExtra("uno", UnitNo);
                startActivity(intentBus);
            }
        });

        gas.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentgas = new Intent(Content.this, GAS.class);
                intentgas.putExtra("uno", UnitNo);
                startActivity(intentgas);
            }
        });

        alarm.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentala = new Intent(Content.this, Alarm.class);
                intentala.putExtra("uno", UnitNo);
                startActivity(intentala);
            }
        });

        sensor.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentsen = new Intent(Content.this, Sensor.class);
                intentsen.putExtra("uno", UnitNo);
                startActivity(intentsen);
            }
        });

        cable.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentcab = new Intent(Content.this, Cable.class);
                intentcab.putExtra("uno", UnitNo);
                startActivity(intentcab);
            }
        });

        geolog.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentgeo = new Intent(Content.this, Geology.class);
                intentgeo.putExtra("uno", UnitNo);
                startActivity(intentgeo);
            }
        });

        riginf.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentrig = new Intent(Content.this, RigInfo.class);
                intentrig.putExtra("uno", UnitNo);
                startActivity(intentrig);
            }
        });

        test.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intenttst = new Intent(Content.this, Testing.class);
                intenttst.putExtra("uno", UnitNo);
                startActivity(intenttst);
            }
        });

        safety.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentasf = new Intent(Content.this, Safety.class);
                intentasf.putExtra("uno", UnitNo);
                startActivity(intentasf);
            }
        });

        tools.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intenttol = new Intent(Content.this, Tools.class);
                intenttol.putExtra("uno", UnitNo);
                startActivity(intenttol);
            }
        });

        stationary.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentsta = new Intent(Content.this, Stationary.class);
                intentsta.putExtra("uno", UnitNo);
                startActivity(intentsta);
            }
        });

        accesory.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentacc = new Intent(Content.this, Accessory.class);
                intentacc.putExtra("uno", UnitNo);
                startActivity(intentacc);
            }
        });
        description.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentdes = new Intent(Content.this, Description.class);
                intentdes.putExtra("uno", UnitNo);
                startActivity(intentdes);
            }
        });

        shortage.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                Intent intentsho = new Intent(Content.this, Shortage.class);
                intentsho.putExtra("uno", UnitNo);
                startActivity(intentsho);
            }
        });


        sharing.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                ArrayList<Uri> uris = new ArrayList<Uri>();
                filename1="Inventory List unit "+UnitNo+".xls";
                filename2= "Unit "+UnitNo+" Shortage.xls";
                File filelocation1 = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename1);
                File filelocation2 = new File(android.os.Environment.getExternalStorageDirectory() + "/GEO-data/Inventory/", filename2);
//
//              Uri path1 = Uri.fromFile(filelocation1);
                Uri path1 = FileProvider.getUriForFile(context, context.getApplicationContext().getPackageName()+".provider",  filelocation1 );
//                Uri path2 = Uri.fromFile(filelocation2);
                Uri path2 = FileProvider.getUriForFile(context, context.getApplicationContext().getPackageName()+".provider",  filelocation2 );
                uris.add(0,path1);
                uris.add(1, path2);
                Intent emailIntent = new Intent(Intent.ACTION_SEND_MULTIPLE);
                emailIntent .setType("vnd.android.cursor.dir/email");
                String to[] = {"amir.arjmand@yahoo.com"};
                emailIntent .putExtra(Intent.EXTRA_EMAIL, to);
                emailIntent.putParcelableArrayListExtra(Intent.EXTRA_STREAM, uris);
                emailIntent .putExtra(Intent.EXTRA_SUBJECT, "Inventory List Unit "+UnitNo);
                startActivity(Intent.createChooser(emailIntent , "Send file Via..."));

            }
        });


//        if (ContextCompat.checkSelfPermission(Content.this, Manifest.permission.WRITE_EXTERNAL_STORAGE) != PackageManager.PERMISSION_GRANTED) {
//            if (ActivityCompat.shouldShowRequestPermissionRationale(Content.this, Manifest.permission.WRITE_EXTERNAL_STORAGE)) {
//                ActivityCompat.requestPermissions(Content.this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, MY_PERMISSION_REQUEST_STORAGE);
//            } else {
//                ActivityCompat.requestPermissions(Content.this, new String[]{Manifest.permission.WRITE_EXTERNAL_STORAGE}, MY_PERMISSION_REQUEST_STORAGE);
//            }
//        } else {
//        }


    }

//    @Override
//    public void onRequestPermissionsResult(int requestCode, @NonNull String[] permissions, @NonNull int[] grantResults) {
//        switch (requestCode) {
//            case MY_PERMISSION_REQUEST_STORAGE: {
//                if (grantResults.length > 0 && grantResults[0] == PackageManager.PERMISSION_GRANTED) {
//                    if (ContextCompat.checkSelfPermission(Content.this, Manifest.permission.WRITE_EXTERNAL_STORAGE) == PackageManager.PERMISSION_GRANTED) {
//                    }
//                } else {
//                    Toast.makeText(this, "No Permission Granted!", Toast.LENGTH_SHORT).show();
//                }
//            }
//
//        }
//
//    }


    public void expo(View view) {

        startActivity(new Intent(Content.this, Unit_List.class));
    }

}







