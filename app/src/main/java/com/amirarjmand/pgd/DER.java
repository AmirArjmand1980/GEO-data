package com.amirarjmand.pgd;

import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;

import java.util.Calendar;
import android.app.Activity;
import android.app.DatePickerDialog;
import android.app.Dialog;
import android.os.Bundle;
import android.view.View;
import android.view.View.OnClickListener;
import android.widget.Button;
import android.widget.DatePicker;
import android.widget.ImageView;
import android.widget.TextView;

public class DER extends AppCompatActivity {

    private TextView  datedisplay,today;
    private DatePicker date_picker;
    private ImageView cal;

    private int year;
    private int month;
    private int day;

    static final int DATE_DIALOG_ID = 100;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(DER.this);
        setContentView(R.layout.activity_der);

        datedisplay=findViewById(R.id.DateDisplay);
        cal=findViewById(R.id.Cal);
        today=findViewById(R.id.Today);

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


                // set current date into textview
                datedisplay.setText(new StringBuilder()
                        // Month is 0 based, so you have to add 1
                        .append(month + 1).append("-")
                        .append(day).append("-")
                        .append(year).append(" "));

                // set current date into Date Picker
                date_picker.init(year, month, day, null);
    }

    public void addButtonListener() {
        final Calendar calendar1 = Calendar.getInstance();

        year = calendar1.get(Calendar.YEAR);
        month = calendar1.get(Calendar.MONTH);
        day = calendar1.get(Calendar.DAY_OF_MONTH);


        // set current date into textview
        datedisplay.setText(new StringBuilder()
                // Month is 0 based, so you have to add 1
                .append(month + 1).append("-")
                .append(day).append("-")
                .append(year).append(" "));

        // set current date into Date Picker
        date_picker.init(year, month, day, null);
                showDialog(DATE_DIALOG_ID);
    }
    @Override
    protected Dialog onCreateDialog(int id) {             

        switch (id) {
            case DATE_DIALOG_ID:
                // set date picker as current date
               DatePickerDialog datePickerDialog=new DatePickerDialog(this, datePickerListener, year, month,day);
               
                return datePickerDialog;
        }
        return null;
    }

    private DatePickerDialog.OnDateSetListener datePickerListener = new DatePickerDialog.OnDateSetListener() {

        // when dialog box is closed, below method will be called.
        public void onDateSet(DatePicker view, int selectedYear,int selectedMonth, int selectedDay) {
            year = selectedYear;
            month = selectedMonth;
            day = selectedDay;

            // set selected date into Text View
            datedisplay.setText(new StringBuilder().append(month + 1)
                    .append("-").append(day).append("-").append(year).append(" "));

            // set selected date into Date Picker
            date_picker.init(year, month, day, null);

        }
    };

}
