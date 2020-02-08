package com.amirarjmand.pgd;

import android.support.design.widget.TabLayout;
import android.support.v4.view.ViewPager;
import android.support.v7.app.AppCompatActivity;
import android.os.Bundle;
import android.support.v7.widget.Toolbar;

import com.amirarjmand.pgd.Inspection_Frags.ViewPagerAdapter;

public class Inspection extends AppCompatActivity {

     private Toolbar toolbar;
     private ViewPager viewPager;
     private ViewPagerAdapter adapter;
     private TabLayout tabLayout;


    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        noTitle noTitle=new noTitle(Inspection.this);
        setContentView(R.layout.activity_inspection);

        toolbar=findViewById(R.id.toolBar);
        setSupportActionBar(toolbar);

        viewPager=findViewById(R.id.pager);
        adapter=new ViewPagerAdapter(getSupportFragmentManager());
        viewPager.setAdapter(adapter);
        tabLayout=findViewById(R.id.tableLayout);


        tabLayout.setupWithViewPager(viewPager);
        
        

    }
    }
                                                      
