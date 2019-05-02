package conversion;
import java.awt.Dimension;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.IOException;









import org.apache.poi.xslf.usermodel.XMLSlideShow;
import org.apache.poi.xslf.usermodel.XSLFSlide;

import com.itextpdf.awt.PdfGraphics2D;
import com.itextpdf.text.Document;
import com.itextpdf.text.DocumentException;
import com.itextpdf.text.Rectangle;
import com.itextpdf.text.pdf.PdfWriter;


public class PPTtoPDFdirectly {



    public static void main(String[] args) throws IOException, DocumentException {

        //load any ppt file
        FileInputStream inputStream = new FileInputStream("a.pptx");
        XMLSlideShow pptx = new XMLSlideShow(inputStream);
        inputStream.close();
        Dimension pgsize = pptx.getPageSize();


        //take first slide and draw it directly into PDF via awt.Graphics2D interface.
        XSLFSlide slide = pptx.getSlides().get(0);

        Document document = new Document();
        PdfWriter pdfWriter = PdfWriter.getInstance(document, new FileOutputStream("a.pdf"));
        document.setPageSize(new Rectangle((float)pgsize.getWidth(), (float) pgsize.getHeight()));
        document.open();

        PdfGraphics2D graphics = new PdfGraphics2D(pdfWriter.getDirectContent(), (float)pgsize.getWidth(), (float)pgsize.getHeight());
        slide.draw(graphics);
        graphics.dispose();

        document.close();
        pptx.close();
    }
}