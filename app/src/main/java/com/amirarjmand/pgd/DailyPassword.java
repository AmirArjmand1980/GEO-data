package com.amirarjmand.pgd;

import android.support.v7.app.AlertDialog;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.view.View;
import android.widget.Button;
import android.widget.EditText;
import android.widget.TextView;

import java.text.SimpleDateFormat;
import java.util.Calendar;
import java.util.Date;
import java.util.Locale;

public class DailyPassword extends AppCompatActivity {
    EditText  c3, c4, c5;
    String datea;
    int pass;
    Button gen;
    TextView passShow, c1, c2;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle = new noTitle(DailyPassword.this);
        setContentView(R.layout.activity_daily_password);

        passShow=findViewById(R.id.Show);
        gen=findViewById(R.id.PassGen);
        c1 = findViewById(R.id.C01);
        c2 = findViewById(R.id.C02);
        c3 = findViewById(R.id.C03);
        c4 = findViewById(R.id.C04);
        c5 = findViewById(R.id.C05);


        SimpleDateFormat sdf = new SimpleDateFormat("dd_MM_yyyy", Locale.getDefault());
        SimpleDateFormat simpleDateFormat = new SimpleDateFormat("dd.MM.yyyy");


        final String currentDateandTime = sdf.format(new Date());
        datea = simpleDateFormat.format(new Date());


        int ye = Calendar.getInstance().get(Calendar.YEAR);
        String yo = String.valueOf(ye);
        yo = yo.substring(3);
        int y = Integer.parseInt(yo);

        int mo = Calendar.getInstance().get(Calendar.MONTH);
        mo = mo + 1;

        int dayy = Calendar.getInstance().get(Calendar.DAY_OF_MONTH);
        TodayDailyPassGen(dayy, mo, y);
        c1.setText(datea);

        gen.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                String daay=c3.getText().toString().trim();
                String moont=c4.getText().toString().trim();
                String yeea=c5.getText().toString().trim();

                if (daay.length()==0 || moont.length()==0 || yeea.length()==0){
                    AlertDialog.Builder adb=new AlertDialog.Builder(DailyPassword.this);
                    adb.setTitle("Note:");
                    adb.setMessage("You have to enter Day, Month and Year values.");
                    adb.setCancelable(true);
                    adb.show();
                }                         else {
                    int day = Integer.parseInt(daay);

                    int mont = Integer.parseInt(moont);

                    yeea = yeea.substring(3);
                    int year = Integer.parseInt(yeea);

                    DailyPassGen(day, mont, year);
                }

            }
        });


    }


    private void TodayDailyPassGen(int day, int month, int year) {

        int op1 = (day * 1000) + (month * 10) + year;
        double op2 = op1 * 1.001;
        double op3 = (op2 / 180) * 3.14159265358979;
        double op4 = Math.sin(op3);
        double op5 = op4 * 10000;
        int pass = (int) op5;
        pass = Math.abs(pass);
        if (pass < 1000) {
            String result = String.valueOf(pass);
            result = "0" + result;
            c2.setText(result);
        } else {
            String result = String.valueOf(pass);
            c2.setText(result);
        }
  

    }

    private void DailyPassGen(int day, int month, int year) {

        int op1 = (day * 1000) + (month * 10) + year;
        double op2 = op1 * 1.001;
        double op3 = (op2 / 180) * 3.14159265358979;
        double op4 = Math.sin(op3);
        double op5 = op4 * 10000;
        int pass = (int) op5;
        pass = Math.abs(pass);
        if (pass < 1000) {
            String result = String.valueOf(pass);
            result = "0" + result;
            passShow.setText(result);
        } else {
            String result = String.valueOf(pass);
            passShow.setText(result);
        }


    }
}

