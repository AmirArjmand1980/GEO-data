package com.amirarjmand.pgd.Inspection_Frags;


import android.os.Bundle;
import android.support.v4.app.Fragment;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.Spinner;
import android.widget.TextView;

import com.amirarjmand.pgd.R;

public class Sensor_Fragment extends Fragment {



    

    @Override
    public View onCreateView(LayoutInflater inflater, ViewGroup container, Bundle savedInstanceState) {
        View v = inflater.inflate(R.layout.fragment_sensor, container, false);
        return v;
    }




    public static Sensor_Fragment newInstance(String text) {

        Sensor_Fragment f = new Sensor_Fragment();
        Bundle b = new Bundle();
        b.putString("msg", text);
        f.setArguments(b);
        return f;
    }



}
