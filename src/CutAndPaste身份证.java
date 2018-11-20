/**
 * Example written by Bruno Lowagie in answer to
 * http://stackoverflow.com/questions/28368317/itext-or-itextsharp-move-text-in-an-existing-pdf
 */

 //鍙傝�冮摼鎺ャ��
//https://developers.itextpdf.com/examples/stamping-content-existing-pdfs-itext5/cut-and-paste-content-page

import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfContentByte;
import com.itextpdf.text.pdf.PdfImportedPage;
import com.itextpdf.text.pdf.PdfReader;
import com.itextpdf.text.pdf.PdfTemplate;
import com.itextpdf.text.pdf.PdfWriter;
import java.io.FileOutputStream;
import java.io.IOException;

public class CutAndPaste身份证 {
 
    /** The original PDF file. */
    public static final String SRC = "学生证2.pdf";
 
    /** The resulting PDF file. */
    public static final String DEST = "generateFile.pdf";
 
    /**
     * Manipulates a PDF file src with the file dest as result
     * @param src the original PDF
     * @param dest the resulting PDF
     * @throws IOException
     * @throws DocumentException
     */
    public void manipulatePdf(String src, String dest)
        throws IOException, DocumentException {
    	// Creating a reader
        PdfReader reader = new PdfReader(src);
        // step 1
        Rectangle pageSize = reader.getPageSize(1);
        
      //  Rectangle toMove = new Rectangle(100, 500, 200, 600);
        //Rectangle toMove = new Rectangle(500, 500, 200, 600);
        
        Document document = new Document(pageSize);
        // step 2
        
        PdfWriter writer  = PdfWriter.getInstance(document, new FileOutputStream(dest));
        // step 3
        document.open();
        // step 4
        
        PdfImportedPage pageUp = writer.getImportedPage(reader, 2);

        PdfImportedPage pageDown = writer.getImportedPage(reader, 1);
        
        PdfContentByte cb = writer.getDirectContent();
        
        PdfTemplate template1 = cb.createTemplate(pageSize.getWidth(), pageSize.getHeight());
        
        template1.rectangle(0, pageSize.getHeight()/2, pageSize.getWidth(), pageSize.getHeight()/2);
     //   template1.rectangle(toMove.getLeft(), toMove.getBottom(),
       //         toMove.getWidth(), toMove.getHeight());
        template1.eoClip();
        template1.newPath();
        template1.addTemplate(pageDown, 0, 0);
        
        
        PdfTemplate template2 = cb.createTemplate(pageSize.getWidth(), pageSize.getHeight());
        
    //    https://itextsupport.com/apidocs/itext5/latest/
        
        template2.rectangle(0, 0, pageSize.getWidth(), pageSize.getHeight()/2);
        
       // template2.rectangle(toMove.getLeft(), toMove.getBottom(),
          //      toMove.getWidth(), toMove.getHeight());
 
        template2.clip();
        template2.newPath();
        
        template2.addTemplate(pageUp, 0, 0);
        
        cb.addTemplate(template1, 0, 0);
        cb.addTemplate(template2, 0, 0);
        
        
        // step 4
        document.close();
        reader.close();
    }
 
    /**
     * Main method.
     * @param    args    no arguments needed
     * @throws DocumentException 
     * @throws IOException
     */
    public static void main(String[] args)
        throws IOException, DocumentException {
      //  File file = new File(DEST);
        //file.getParentFile().mkdirs();
        new CutAndPaste身份证().manipulatePdf(SRC, DEST);
    }
}