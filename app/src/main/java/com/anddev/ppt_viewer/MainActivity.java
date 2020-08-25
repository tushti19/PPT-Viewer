package com.anddev.ppt_viewer;

import androidx.annotation.Nullable;
import androidx.annotation.RequiresApi;
import androidx.appcompat.app.AppCompatActivity;
import androidx.core.content.res.ResourcesCompat;

import android.annotation.SuppressLint;
import android.content.Intent;
import android.content.res.Resources;
import android.graphics.Bitmap;
import android.graphics.BitmapFactory;
import android.graphics.Color;
import android.graphics.Typeface;
import android.os.Build;
import android.os.Bundle;
import android.util.Log;
import android.util.TypedValue;
import android.view.View;
import android.widget.Button;
import android.widget.ImageView;
import android.widget.LinearLayout;
import android.widget.TextView;

import org.apache.poi.sl.draw.DrawPaint;
import org.apache.poi.sl.usermodel.Line;
import org.apache.poi.sl.usermodel.PaintStyle;
import org.apache.poi.sl.usermodel.PlaceableShape;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFAutoShape;
import org.apache.poi.xslf.usermodel.XSLFBackground;
import org.apache.poi.xslf.usermodel.XSLFColor;
import org.apache.poi.xslf.usermodel.XSLFConnectorShape;
import org.apache.poi.xslf.usermodel.XSLFPictureData;
import org.apache.poi.xslf.usermodel.XSLFPictureShape;
import org.apache.poi.xslf.usermodel.XSLFShape;
import org.apache.poi.xslf.usermodel.XSLFSlide;
import org.apache.poi.xslf.usermodel.XSLFSlideLayout;
import org.apache.poi.xslf.usermodel.XSLFSlideMaster;
import org.apache.poi.xslf.usermodel.XSLFTextParagraph;
import org.apache.poi.xslf.usermodel.XSLFTextRun;
import org.apache.poi.xslf.usermodel.XSLFTextShape;
import org.apache.poi.xslf.usermodel.XSLFTheme;
import org.apache.xmlbeans.XmlObject;
import org.apache.xmlbeans.XmlString;
import org.openxmlformats.schemas.drawingml.x2006.main.CTBaseStyles;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontCollection;
import org.openxmlformats.schemas.drawingml.x2006.main.CTFontScheme;
import org.openxmlformats.schemas.drawingml.x2006.main.CTOfficeStyleSheet;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextBody;
import org.openxmlformats.schemas.drawingml.x2006.main.CTTextFont;
import org.openxmlformats.schemas.presentationml.x2006.main.CTBackground;
import org.openxmlformats.schemas.presentationml.x2006.main.CTSlide;

import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.IOException;
import java.util.List;

public class MainActivity extends AppCompatActivity {

    Button btn;
    LinearLayout linearLayout;

    Intent intent;

    @Override
    protected void onCreate(Bundle savedInstanceState) {
        super.onCreate(savedInstanceState);
        setContentView(R.layout.activity_main);
        System.setProperty("org.apache.poi.javax.xml.stream.XMLInputFactory", "com.fasterxml.aalto.stax.InputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLOutputFactory", "com.fasterxml.aalto.stax.OutputFactoryImpl");
        System.setProperty("org.apache.poi.javax.xml.stream.XMLEventFactory", "com.fasterxml.aalto.stax.EventFactoryImpl");

        btn = findViewById(R.id.Button);
        linearLayout = findViewById(R.id.LinearLayout);


        btn.setOnClickListener(new View.OnClickListener() {
            @Override
            public void onClick(View v) {
                intent = new Intent(Intent.ACTION_GET_CONTENT);
                intent.setType("*/*");
                startActivityForResult(intent, 10);
                btn.setVisibility(View.INVISIBLE);
            }
        });


    }


    @Override
    protected void onActivityResult(int requestCode, int resultCode, @Nullable Intent data) {
        super.onActivityResult(requestCode, resultCode, data);

        try {
            if (resultCode == RESULT_OK) {
                switch (requestCode) {
                    case 10:
                        //this is action performed after openDocumentFromFileManager() when doc is selected
                        FileInputStream inputStream = (FileInputStream) getContentResolver().openInputStream(data.getData());
                        XMLSlideShow ppt = new XMLSlideShow(inputStream);



                       // readImages(ppt);

                       // readAllText(ppt);
                        readText(ppt);




                }
            }

        } catch (FileNotFoundException e) {
            e.printStackTrace();
        } catch (IOException e) {
            e.printStackTrace();
        }


    }

    private void readAllText(XMLSlideShow ppt) {


        for (XSLFSlide slide : ppt.getSlides()) {
            CTSlide ctSlide = slide.getXmlObject();
            XmlObject[] allText = ctSlide.selectPath(
                    "declare namespace a='http://schemas.openxmlformats.org/drawingml/2006/main' " +
                            ".//a:t"
            );
            for (int i = 0; i < allText.length; i++) {
                if (allText[i] instanceof XmlString) {
                    XmlString xmlString = (XmlString)allText[i];
                    String text = xmlString.getStringValue();
                    System.out.println(text);


                   addTextViews(text);
                }
                else
                {
                    Log.d("picture","pic");
                }
            }
        }



    }

    private void addTextViews(String text) {

        TextView textView = new TextView(this);
        LinearLayout.LayoutParams layoutParams = new LinearLayout.LayoutParams(LinearLayout.LayoutParams.WRAP_CONTENT, LinearLayout.LayoutParams.WRAP_CONTENT);
        layoutParams.height = LinearLayout.LayoutParams.WRAP_CONTENT;
        layoutParams.width = LinearLayout.LayoutParams.WRAP_CONTENT;
        textView.setTextSize(20f);
        textView.setTextColor(Color.BLACK);
        textView.setText(text);
        textView.setLayoutParams(layoutParams);
        linearLayout.addView(textView);

    }


    private void readText(XMLSlideShow ppt) {


        for(XSLFSlideMaster master : ppt.getSlideMasters()) {
            for (XSLFSlideLayout layout : master.getSlideLayouts()) {
                System.out.println(layout.getType());
            }
        }


        XSLFSlide[] slide2 = ppt.getSlides().toArray(new XSLFSlide[1]);
        for (int i = 0; i < slide2.length; i++)
            {
                addDivider();
                XSLFTheme theme = slide2[i].getTheme();
                CTOfficeStyleSheet styleSheet = theme.getXmlObject();
                CTBaseStyles themeElements = styleSheet.getThemeElements();
                CTFontScheme fontScheme = themeElements.getFontScheme();
                CTFontCollection fontCollection = fontScheme.getMajorFont();
                CTFontCollection minorFont = fontScheme.getMinorFont();
                CTTextFont latin = minorFont.getLatin();
                CTTextFont arial = minorFont.getCs();
                String majorFontName = theme.getMajorFont();
                String minorFontName = theme.getMinorFont();
                Log.d("theme", String.valueOf(theme));
               // Log.d("stylesheet", String.valueOf(styleSheet));
              //  Log.d("themeElements", String.valueOf(themeElements));
              //  Log.d("fontScheme", String.valueOf(fontScheme));
             //   Log.d("fontCollection", String.valueOf(fontCollection));
                Log.d("font", majorFontName + " " + minorFontName);

                XSLFBackground bg = slide2[i].getBackground();
                org.apache.poi.java.awt.Color f = bg.getFillColor();
                int AlphaBg = f.getAlpha();
                int RedBg = f.getRed();
                int BlueBg = f.getBlue();
                int GreenBg = f.getGreen();
                int ColorBg = Color.argb(AlphaBg,RedBg,GreenBg,BlueBg);
                //linearLayout.setBackgroundColor(ColorBg);




                Log.d("Bg", "*"+ AlphaBg + BlueBg + RedBg + GreenBg);












                List<XSLFShape> shapes = slide2[i].getShapes();


                for(int j =0;j<shapes.size();j++){

                    Log.d("shapename", shapes.get(j).getShapeName());

                    if(shapes.get(j) instanceof XSLFTextShape) {

                        Log.d("Text", "yes");

                        XSLFTextShape txShape = (XSLFTextShape) shapes.get(j);
                        Log.d("textshape", txShape.getShapeName());
                        for (XSLFTextParagraph xslfParagraph : txShape.getTextParagraphs()) {
                            //System.out.println(xslfParagraph.getText());




                            PaintStyle fontColor = null;
                            Color color = null;
                            String fontFamily = " ";
                            Double fontSize = 0.0;
                            boolean italic = false;
                            boolean bold = false;
                            boolean underline = false;
                            int COLOR = 0;


                            for(XSLFTextRun text : xslfParagraph.getTextRuns() ){


                                fontColor = text.getFontColor();
                                Log.d("Color", String.valueOf(fontColor));
                                fontFamily = text.getFontFamily();
                                Log.d("Font Family",fontFamily);
                                fontSize = text.getFontSize();
                                italic = text.isItalic();
                                bold = text.isBold();
                                underline = text.isUnderlined();


                                PaintStyle.SolidPaint sp = (PaintStyle.SolidPaint) fontColor;

                                org.apache.poi.java.awt.Color c = DrawPaint.applyColorTransform(sp.getSolidColor());
                                int A = c.getAlpha();
                                int B = c.getBlue();
                                int R = c.getRed();
                                int G = c.getGreen();

                                Log.d("hash", "*"+ A + B + R + G);

                                 COLOR = Color.argb(A,R,G,B);





                            }





                            TextView textView = new TextView(this);
                            LinearLayout.LayoutParams layoutParams = new LinearLayout.LayoutParams(LinearLayout.LayoutParams.WRAP_CONTENT, LinearLayout.LayoutParams.WRAP_CONTENT);
                            layoutParams.height = LinearLayout.LayoutParams.WRAP_CONTENT;
                            layoutParams.width = LinearLayout.LayoutParams.WRAP_CONTENT;
                            textView.setTextSize(fontSize.floatValue());
                            textView.setTextColor(COLOR);



                            //Typeface typeface = ResourcesCompat.getFont(this, R.font.times_new_roman);
                           // textView.setTypeface(typeface);


                            if(bold) {textView.setTypeface(Typeface.defaultFromStyle(Typeface.BOLD));}
                            if(italic){textView.setTypeface(Typeface.defaultFromStyle(Typeface.ITALIC));}

                            textView.setText(xslfParagraph.getText());
                            textView.setLayoutParams(layoutParams);
                            linearLayout.addView(textView);
                        }
                    }
                    if(shapes.get(j) instanceof XSLFPictureShape) {
                        Log.d("Picture", "yes");

                        XSLFPictureShape picShape = (XSLFPictureShape) shapes.get(j);
                        XSLFPictureData data = picShape.getPictureData();

                        byte[] bytes = data.getData();
                        String fileName = data.getFileName();

                        ImageView imageView = new ImageView(MainActivity.this);
                        LinearLayout.LayoutParams params = new LinearLayout.LayoutParams(LinearLayout.LayoutParams.WRAP_CONTENT, LinearLayout.LayoutParams.WRAP_CONTENT);
                        params.width = LinearLayout.LayoutParams.WRAP_CONTENT;

                        Resources r = getResources();
                        int margin = (int) TypedValue.applyDimension(
                                TypedValue.COMPLEX_UNIT_DIP,
                                200,
                                r.getDisplayMetrics());

                        params.height = margin;


                        Bitmap bmp = BitmapFactory.decodeByteArray(bytes, 0, bytes.length);

                        // Set the Bitmap data to the ImageView
                        imageView.setImageBitmap(bmp);

                        linearLayout.addView(imageView);




                        //readImages(ppt);
                    }

                    if(shapes.get(j) instanceof  XSLFConnectorShape)
                    {
                        XSLFConnectorShape line = (XSLFConnectorShape) shapes.get(j);
                        Log.d("Line","yes" + line.getShapeName() );
                    }


                   if(shapes.get(j) instanceof XSLFAutoShape) {
                       Log.d("autoshape", "yes");
                       XSLFAutoShape autoShape = (XSLFAutoShape) shapes.get(j);
                       Log.d("autoshape", autoShape.getShapeName());

                   }

                }

            }

            /*try {
                XSLFNotes mynotes = slide2[i].getNotes();
                for (XSLFShape shape : mynotes) {
                    Log.d("HII", "Hii");
                    if (shape instanceof XSLFTextShape) {
                        XSLFTextShape txShape = (XSLFTextShape) shape;
                        for (XSLFTextParagraph xslfParagraph : txShape.getTextParagraphs()) {

                            Log.d("HII 1", "Hii");
                            System.out.println(xslfParagraph.getText());
                        }
                    }
                }


            } catch (Exception e) {

            }*/



    }

    private void addDivider() {


        int dividerHeight = (int) (getResources().getDisplayMetrics().density * 1);
        // 1dp to pixels
        int ht = (int) (getResources().getDisplayMetrics().density * 20);


        View view = new View(this);
        LinearLayout.LayoutParams layoutParams = new LinearLayout.LayoutParams(LinearLayout.LayoutParams.MATCH_PARENT, dividerHeight);
        layoutParams.topMargin = ht;
        layoutParams.bottomMargin = ht;
        view.setLayoutParams(layoutParams);
        view.setBackgroundColor(Color.BLACK);
        linearLayout.addView(view);


    }

    private void readImages(XMLSlideShow ppt) {

        btn.setVisibility(View.INVISIBLE);

        for (XSLFPictureData data : ppt.getPictureData()) {
            byte[] bytes = data.getData();
            String fileName = data.getFileName();
           // Log.d("File name", " here " + fileName);
           // Log.d("No. of bytes", "is " + bytes);


            ImageView imageView = new ImageView(MainActivity.this);
            LinearLayout.LayoutParams params = new LinearLayout.LayoutParams(LinearLayout.LayoutParams.WRAP_CONTENT, LinearLayout.LayoutParams.WRAP_CONTENT);
            params.width = LinearLayout.LayoutParams.WRAP_CONTENT;

            Resources r = getResources();
            int margin = (int) TypedValue.applyDimension(
                    TypedValue.COMPLEX_UNIT_DIP,
                    200,
                    r.getDisplayMetrics());

            params.height = margin;


            Bitmap bmp = BitmapFactory.decodeByteArray(bytes, 0, bytes.length);

            // Set the Bitmap data to the ImageView
            imageView.setImageBitmap(bmp);

            linearLayout.addView(imageView);


        }


    }







    private void readShapes(XMLSlideShow ppt) {

        for (XSLFSlide slide : ppt.getSlides()) {
            for (XSLFShape sh : slide.getShapes()) {
                // name of the shape
                String name = sh.getShapeName();
                // shapes's anchor which defines the position of this shape in the slide
                if (sh instanceof PlaceableShape) {
                    //  java.awt.geom.Rectangle2D anchor = ((PlaceableShape)sh).getAnchor();
                }
                if (sh instanceof XSLFConnectorShape) {
                    XSLFConnectorShape line = (XSLFConnectorShape) sh;
                    // work with Line
                } else if (sh instanceof XSLFTextShape) {
                    XSLFTextShape shape = (XSLFTextShape) sh;

                    // work with a shape that can hold text
                } else if (sh instanceof XSLFPictureShape) {
                    XSLFPictureShape shape = (XSLFPictureShape) sh;
                    // work with Picture
                }
            }
        }
    }
}


