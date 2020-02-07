package com.amirarjmand.pgd.Inspection_Frags;

import android.os.Bundle;
import android.support.annotation.Nullable;
import android.support.v4.app.Fragment;
import android.support.v4.app.FragmentManager;
import android.support.v4.app.FragmentPagerAdapter;
import android.widget.Spinner;

public class ViewPagerAdapter extends FragmentPagerAdapter {



    public ViewPagerAdapter(FragmentManager fm) {
        super(fm);
    }

    @Override
    public Fragment getItem(int position) {
        switch(position) {

            case 0: return Sensor_Fragment.newInstance("Sensor Inspection, Instance 1");
            case 1: return Gas_Fragment.newInstance("SecondFragment, Instance 2");
            case 2: return Computer_Fragment.newInstance("ThirdFragment, Instance 3");
            case 3: return Geology_Fragment.newInstance("ThirdFragment, Instance 4");
            default: return Sensor_Fragment.newInstance("ThirdFragment, Default");


    }
    }

    @Override
    public int getCount() {
        return 4;
    }


    @Nullable
    @Override
    public CharSequence getPageTitle(int position) {
        if (position==0){return "Sensors";}
        if (position==1){return "Gas Equipments";}
        if (position==2){return "Computer Equipments";}
        else return "Geology Equipments" ;
    }
}
