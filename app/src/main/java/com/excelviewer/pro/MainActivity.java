package com.excelviewer.pro;

import android.app.Activity;
import android.os.Bundle;
import android.view.Gravity;
import android.widget.LinearLayout;
import android.widget.TextView;
import android.graphics.Color;

public class MainActivity extends Activity {

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        LinearLayout layout = new LinearLayout(this);
        layout.setOrientation(LinearLayout.VERTICAL);
        layout.setGravity(Gravity.CENTER);
        layout.setPadding(30, 30, 30, 30);
        layout.setBackgroundColor(Color.WHITE);

        TextView title = new TextView(this);
        title.setText("Excel Viewer Pro");
        title.setTextSize(28);
        title.setTextColor(Color.BLACK);
        title.setGravity(Gravity.CENTER);

        TextView message = new TextView(this);
        message.setText("Поиск по Excel-файлам\nXLS и XLSX");
        message.setTextSize(18);
        message.setTextColor(Color.DKGRAY);
        message.setGravity(Gravity.CENTER);
        message.setPadding(0, 30, 0, 0);

        layout.addView(title);
        layout.addView(message);

        setContentView(layout);
    }
}
