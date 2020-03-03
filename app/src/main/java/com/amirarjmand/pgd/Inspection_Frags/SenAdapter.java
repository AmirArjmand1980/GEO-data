package com.amirarjmand.pgd.Inspection_Frags;

import android.content.Context;
import android.support.annotation.NonNull;
import android.support.v7.widget.RecyclerView;
import android.text.Editable;
import android.text.TextWatcher;
import android.view.LayoutInflater;
import android.view.View;
import android.view.ViewGroup;
import android.widget.EditText;
import android.widget.TextView;

import com.amirarjmand.pgd.R;

import java.util.ArrayList;

public class SenAdapter extends RecyclerView.Adapter<SenAdapter.SenViewHolder> {

    Context context;
    ArrayList<SenRecData_Model> senRecData_models;

    public SenAdapter(Context context, ArrayList<SenRecData_Model> senRecData_models, ArrayList<SenRecHeader_Model> senRecHeader_models) {
        this.context = context;
        this.senRecData_models = senRecData_models;
        this.senRecHeader_models = senRecHeader_models;
    }

    ArrayList<SenRecHeader_Model>senRecHeader_models;


    @NonNull
    @Override
    public SenViewHolder onCreateViewHolder(@NonNull ViewGroup viewGroup, int i) {
        LayoutInflater layoutInflater=LayoutInflater.from(context);
        View view=layoutInflater.inflate(R.layout.inspection_recycler_item, null);
        return new SenViewHolder(view) ;
    }

    @Override
    public void onBindViewHolder(@NonNull SenViewHolder holder, int position) {
        holder.header.setText(senRecHeader_models.get(position).getSenRecHeader());
        holder.nws.setText(senRecData_models.get(position).getQtyUseSensor());
        holder.nss.setText(senRecData_models.get(position).getQtySpareSensor());
        holder.st.setText(senRecData_models.get(position).getSenType());

    }

    @Override
    public int getItemCount() {
        return senRecData_models.size();
    }

    public class   SenViewHolder extends RecyclerView.ViewHolder{

        EditText nws, nss, st;
        TextView header;

        public SenViewHolder(@NonNull View itemView) {
            super(itemView);
            nws=itemView.findViewById(R.id.NoWorkSen);
            nss=itemView.findViewById(R.id.NoSprSen);
            st=itemView.findViewById(R.id.SenType);
            header=itemView.findViewById(R.id.Sensor_name);


            nws.addTextChangedListener(new TextWatcher() {
                @Override
                public void beforeTextChanged(CharSequence s, int start, int count, int after) {
                    
                }

                @Override
                public void onTextChanged(CharSequence s, int start, int before, int count) {
                    senRecData_models.get(getAdapterPosition()).setQtyUseSensor(nws.getText().toString());

                }

                @Override
                public void afterTextChanged(Editable s) {

                }
            });

            nss.addTextChangedListener(new TextWatcher() {
                @Override
                public void beforeTextChanged(CharSequence s, int start, int count, int after) {

                }

                @Override
                public void onTextChanged(CharSequence s, int start, int before, int count) {
                    senRecData_models.get(getAdapterPosition()).setQtyUseSensor(nss.getText().toString());

                }

                @Override
                public void afterTextChanged(Editable s) {

                }
            });

            st.addTextChangedListener(new TextWatcher() {
                @Override
                public void beforeTextChanged(CharSequence s, int start, int count, int after) {

                }

                @Override
                public void onTextChanged(CharSequence s, int start, int before, int count) {
                    senRecData_models.get(getAdapterPosition()).setQtyUseSensor(st.getText().toString());

                }

                @Override
                public void afterTextChanged(Editable s) {

                }
            });


        }
    }
}
