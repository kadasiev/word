package org.example;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;

import java.io.FileOutputStream;
import java.io.IOException;

public class CreateWordDocument {
  public static void main(String[] args) {
    // Create a new document
    XWPFDocument document = new XWPFDocument();

    // Create a new paragraph
    XWPFParagraph paragraph = document.createParagraph();
    XWPFRun run = paragraph.createRun();
    run.setText("Hello, this is a sample Word document created using Apache POI!");

    // Save the document to a file
    try (FileOutputStream out = new FileOutputStream("sample-doc.docx")) {
      document.write(out);
      System.out.println("Document created successfully!");
    } catch (IOException e) {
      e.printStackTrace();
    }
  }
}
