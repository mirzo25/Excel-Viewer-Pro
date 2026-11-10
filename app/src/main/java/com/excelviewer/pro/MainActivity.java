package com.excelviewer.pro;

import android.app.Activity;
import android.os.Bundle;
import android.content.Intent;
import android.net.Uri;
import android.graphics.Color;
import android.view.Gravity;
import android.widget.*;
import org.apache.poi.ss.usermodel.*;
import java.io.InputStream;
import java.util.Locale;

public class MainActivity extends Activity {

    private LinearLayout root;
    private LinearLayout tabs;
    private TextView result;
    private EditText search;
    private Workbook workbook;
    private DataFormatter formatter = new DataFormatter();

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);

        root = new LinearLayout(this);
        root.setOrientation(LinearLayout.VERTICAL);
        root.setPadding(12, 12, 12, 12);
        root.setBackgroundColor(Color.WHITE);

        TextView title = new TextView(this);
        title.setText("Excel Viewer Pro");
        title.setTextSize(26);
        title.setTextColor(Color.BLACK);
        title.setGravity(Gravity.CENTER);
        root.addView(title);

        Button open = new Button(this);
        open.setText("ОТКРЫТЬ EXCEL ФАЙЛ");
        root.addView(open);
        open.setOnClickListener(v -> openFile());

        search = new EditText(this);
        search.setHint("Название товара или код");
        root.addView(search);

        Button find = new Button(this);
        find.setText("ПОИСК ПО ВСЕМ ЛИСТАМ");
        root.addView(find);
        find.setOnClickListener(v -> findText());

        HorizontalScrollView horizontal = new HorizontalScrollView(this);
        tabs = new LinearLayout(this);
        tabs.setOrientation(LinearLayout.HORIZONTAL);
        horizontal.addView(tabs);
        root.addView(horizontal);

        ScrollView scroll = new ScrollView(this);
        result = new TextView(this);
        result.setTextColor(Color.BLACK);
        result.setTextSize(14);
        result.setPadding(8, 12, 8, 12);
        scroll.addView(result);

        root.addView(scroll, new LinearLayout.LayoutParams(
                -1, 0, 1));

        result.setText("Выберите Excel-файл XLS или XLSX.");

        setContentView(root);
    }

    private void openFile() {
        Intent intent = new Intent(Intent.ACTION_OPEN_DOCUMENT);
        intent.setType("*/*");
        intent.addCategory(Intent.CATEGORY_OPENABLE);

        intent.putExtra(Intent.EXTRA_MIME_TYPES, new String[]{
            "application/vnd.ms-excel",
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        });

        startActivityForResult(intent, 100);
    }

    @Override
    protected void onActivityResult(int requestCode,
            int resultCode, Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        if (requestCode == 100 && resultCode == RESULT_OK
                && data != null && data.getData() != null) {
            loadFile(data.getData());
        }
    }

    private void loadFile(Uri uri) {
        try {
            if (workbook != null) {
                workbook.close();
            }

            InputStream input =
                    getContentResolver().openInputStream(uri);

            if (input == null) {
                throw new Exception("Файл не удалось открыть");
            }

            try {
                workbook = WorkbookFactory.create(input);
            } finally {
                input.close();
            }

            showTabs();
            showSheet(0);

        } catch (Exception e) {
            result.setText("Ошибка: " + e.getMessage());
        }
    }

    private void showTabs() {
        tabs.removeAllViews();

        if (workbook == null) return;

        for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
            final int number = i;

            Button tab = new Button(this);
            tab.setText(workbook.getSheetName(i));
            tab.setOnClickListener(v -> showSheet(number));

            tabs.addView(tab);
        }
    }
