package com.amirarjmand.pgd.Inspection_Frags;

import android.os.Bundle;
import android.support.v7.app.AppCompatActivity;
import android.support.v7.widget.LinearLayoutManager;
import android.support.v7.widget.RecyclerView;

import com.amirarjmand.pgd.R;
import com.amirarjmand.pgd.noTitle;

import java.util.ArrayList;

public class Inspection_Sensor extends AppCompatActivity {

    RecyclerView senRec;
    SenAdapter SenAdapter;
    ArrayList<SenRecData_Model> senRecData_models;
    ArrayList<SenRecHeader_Model>senRecHeader_models;
    

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(Inspection_Sensor.this);
        setContentView(R.layout.activity_inspection__sensor);

        senRec=findViewById(R.id.SenRecycler);
        senRecData_models=new ArrayList<>();
        senRecHeader_models=new ArrayList<>();

        senRecHeader_models.add(new SenRecHeader_Model("Depth Tracking (B10)"));
        senRecHeader_models.add(new SenRecHeader_Model("Hook Load"));
        senRecHeader_models.add(new SenRecHeader_Model("Stand Pipe Pressure"));
        senRecHeader_models.add(new SenRecHeader_Model("Annular Pressure"));
        senRecHeader_models.add(new SenRecHeader_Model("Pit Level"));
        senRecHeader_models.add(new SenRecHeader_Model("Flow Out"));
        senRecHeader_models.add(new SenRecHeader_Model("Pump Strokes"));
        senRecHeader_models.add(new SenRecHeader_Model("Torque"));
        senRecHeader_models.add(new SenRecHeader_Model("RPM"));
        senRecHeader_models.add(new SenRecHeader_Model("Mud Weight IN"));
        senRecHeader_models.add(new SenRecHeader_Model("Mud Weight OUT"));
        senRecHeader_models.add(new SenRecHeader_Model("Temperature IN"));
        senRecHeader_models.add(new SenRecHeader_Model("Temperature OUT"));
        senRecHeader_models.add(new SenRecHeader_Model("Conductivity IN"));
        senRecHeader_models.add(new SenRecHeader_Model("Conductivity OUT"));





        senRecData_models.add(new SenRecData_Model("1", "1", "Proximity"));
        senRecData_models.add(new SenRecData_Model("1", "0", "Hydraulic"));
        senRecData_models.add(new SenRecData_Model("1", "1", "Hydraulic"));
        senRecData_models.add(new SenRecData_Model("1", "0", "Hydraulic"));
        senRecData_models.add(new SenRecData_Model("8", "1", "Ultra sonic"));
        senRecData_models.add(new SenRecData_Model("1", "0", "Ultra sonic"));
        senRecData_models.add(new SenRecData_Model("3", "3", "Proximity"));
        senRecData_models.add(new SenRecData_Model("0", "0", "Third Party"));
        senRecData_models.add(new SenRecData_Model("0", "0", "Third Party"));
        senRecData_models.add(new SenRecData_Model("1", "0", "Dia.Press"));
        senRecData_models.add(new SenRecData_Model("1", "0", "Dia.Press"));
        senRecData_models.add(new SenRecData_Model("1", "0", "Electrical"));
        senRecData_models.add(new SenRecData_Model("1", "0", "Electrical"));
        senRecData_models.add(new SenRecData_Model("0", "0", "Chemical"));
        senRecData_models.add(new SenRecData_Model("0", "0", "Chemical"));


        SenAdapter=new SenAdapter(Inspection_Sensor.this, senRecData_models, senRecHeader_models);
        senRec.setLayoutManager(new LinearLayoutManager(getApplicationContext(), LinearLayoutManager.HORIZONTAL, false));
        senRec.setAdapter(SenAdapter);


    }
}
