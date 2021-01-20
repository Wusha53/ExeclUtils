package com.wusha.execlutils;

import androidx.appcompat.app.AppCompatActivity;

import android.content.Context;
import android.content.Intent;
import android.inputmethodservice.Keyboard;
import android.net.Uri;
import android.os.Bundle;
import android.view.View;
import android.widget.EditText;

import java.io.File;
import java.io.FileOutputStream;
import java.io.InputStream;
import java.io.OutputStream;
import java.text.SimpleDateFormat;
import java.util.List;

public class MainActivity extends AppCompatActivity {


    EditText output;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        output = (EditText) findViewById(R.id.textOut);
        findViewById(R.id.read_xlsx).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                InputStream stream = getResources().openRawResource(R.raw.test);
                try {

                    printlnToUser(ExeclUtils.ReadExcelUtilsXLSX(stream));
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
        findViewById(R.id.read_lsx).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                InputStream stream = getResources().openRawResource(R.raw.test2);
                try {

                    printlnToUser(ExeclUtils.ReadExcelUtilsXLS(stream));
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
        findViewById(R.id.read_csv).setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {

                InputStream stream = getResources().openRawResource(R.raw.test3);
                try {
                    List<String> a= ExeclUtils.ReadExcelUtilsCSV(stream);
                    printlnToUser(a.toString());
                } catch (Exception e) {
                    e.printStackTrace();
                }
            }
        });
    }

    /**
     * print line to the output TextView
     * @param str
     */
    private void printlnToUser(String str) {
        final String string = str;
        if (output.length()>8000) {
            CharSequence fullOutput = output.getText();
            fullOutput = fullOutput.subSequence(5000,fullOutput.length());
            output.setText(fullOutput);
            output.setSelection(fullOutput.length());
        }
        output.append(string+"\n");
    }




}
