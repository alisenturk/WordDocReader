/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package com.altar.worddocreader;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

/**
 *
 * @author asenturk
 */
public class WordDocReaderMain {

    public static void main(String[] args) {
        HashMap<String, String> keys = new HashMap<String, String>();
        keys.put("#name#", "Ali");
        keys.put("#surname#", "Şentürk");
        keys.put("#birthdate#", "09/08/1980");
        keys.put("#companyName#", "A.B.C. A.Ş. İnsan Kaynakları Müdürlüğü");
       
        readWordDoc("c:/temp/test1.docx","c:/temp/test2.docx", keys);
        
    }

    public static void readWordDoc(String orginalFilePath, String newFilePath, HashMap<String, String> keyMap) {
        FileInputStream fis = null;
        FileOutputStream out = null;
        try {
            fis = new FileInputStream(orginalFilePath);
            XWPFDocument doc = new XWPFDocument(fis);        
            doc = replaceText(doc, keyMap);
        
            out = new FileOutputStream(new File(newFilePath));

            doc.write(out);
        } catch (Exception e) {
            e.printStackTrace();
        } finally {
            if (fis != null) {
                try {
                    if (out != null) {
                        out.close();
                    }
                    fis.close();
                    out = null;
                    fis = null;
                } catch (IOException ioEx) {
                    ioEx.printStackTrace();
                }
            }
        }
    }

    public static XWPFDocument replaceText(XWPFDocument doc, HashMap<String, String> keys) throws Exception {
        String txt = "";
        int txtPosition = 0;
        String key = "";
        String val = "";
        for (XWPFParagraph p : doc.getParagraphs()) { //Dökümandaki her bir paragraf okuması yapılıyor.
            for (XWPFRun run : p.getRuns()) { //paragraf içindeki satırlar okunuyor.
                txtPosition = run.getTextPosition();
                txt = run.getText(txtPosition);
                for (Map.Entry<String, String> entry : keys.entrySet()) { //keymap içinde gönderilen alanlar keymap'teki değerleri ile değiştiriliyor.
                    key = entry.getKey();
                    val = entry.getValue();
                    if (txt != null && txt.indexOf(key) > -1) {
                        txt = txt.replace(key, val);
                        run.setText(txt, 0);
                    }
                }
            }
        }

        return doc; 

    }

}
