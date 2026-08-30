/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package timetable;

import java.io.IOException;  
import java.text.DateFormat;
import org.jsoup.Jsoup;  
import org.jsoup.nodes.Document;
import org.jsoup.select.Elements;
import java.text.SimpleDateFormat;
import java.util.Calendar;
//import org.jsoup.nodes.Element;

import org.apache.poi.xwpf.usermodel.XWPFDocument;
import org.apache.poi.xwpf.usermodel.XWPFParagraph;
import org.apache.poi.xwpf.usermodel.XWPFRun;
import java.text.DateFormatSymbols;
import java.io.FileOutputStream;
import java.net.URISyntaxException;
import java.util.HashMap;
import java.util.Map;
import org.apache.poi.openxml4j.exceptions.InvalidFormatException;
import org.apache.poi.openxml4j.opc.OPCPackage;
import org.apache.poi.xwpf.usermodel.BodyElementType;
import org.apache.poi.xwpf.usermodel.IBodyElement;
import org.apache.poi.xwpf.usermodel.XWPFTable;
import org.apache.poi.xwpf.usermodel.XWPFTableCell;
import org.apache.poi.xwpf.usermodel.XWPFTableRow;
import org.jsoup.nodes.Element;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTR;
import org.openxmlformats.schemas.wordprocessingml.x2006.main.CTRPr;






/**
 *
 * @author ivani
 */
public class Timetable {

    /**
     * @param args the command line arguments
     * @throws java.net.URISyntaxException
     * @throws java.io.IOException
     * @throws org.apache.poi.openxml4j.exceptions.InvalidFormatException
     */
     
    public static void main(String[] args) throws URISyntaxException, IOException, InvalidFormatException {

        String [] date_full = new String [7];

        String [] text = new String [7];
        String [] strong_big = new String [7];
        String [][] strong_matr = new String [7][1000];
        String [][] em_matr = new String [7][1000];
        //System.out.println(day__text);
        //System.out.println("Title : " + gg);
        
        //SimpleDateFormat dateFormat = new SimpleDateFormat("yyyy-MM-dd");
        //Date date = new Date();
        
        Calendar calendar = Calendar.getInstance();//new GregorianCalendar();
        //calendar.format();
        DateFormat df_file = new SimpleDateFormat("dd MMMM yyyy");
        DateFormat df = new SimpleDateFormat("yyyy-MM-dd");
        DateFormat df1= new SimpleDateFormat("d MMMM", myDateFormatSymbols );
        System.out.println(df.format(calendar.getTime()));
        
        int i = calendar.get(Calendar.DAY_OF_WEEK);
        while (i!=2)
        {
            calendar.add(Calendar.DAY_OF_MONTH, +1);
            i = calendar.get(Calendar.DAY_OF_WEEK);
        }
        String Output_file = df_file.format(calendar.getTime());
        //calendar.add(Calendar.DAY_OF_MONTH, +7);

        for (i = 0; i < 7; i++) {
            String date = df.format(calendar.getTime());
            Document doc = Jsoup.connect("https://azbyka.ru/days/"+ date).get();
            Elements day__text = doc.select("div.text.day__text");
            //day__text_strong = doc.select("div.day__text strong");
            //day__text_em = doc.select("div.day__text em");
            Elements p = day__text.select("p");
            Elements ul = day__text.select("ul");
            //strong [i] = day__text_strong.text();
            //Element ptemp = ul.get(0);
            //text [i] = ptemp.text();
            
            //Elements p01 = new Elements();
            //p01.add(ptemp);
            //ptemp = p.get(1);
            //p01.add(ptemp);
            String top_text = doc.select("div.post").text();
            
            if (!p.isEmpty() )
                strong_big[i] = p.get(0).text();
            else
                strong_big[i] = "";
            if (strong_big[i].length()!=0) strong_big[i] = strong_big[i];
            //strong_big[i] = strong_big[i] + p.get(1).text();
            strong_big[i] = top_text + ". " + strong_big[i];
            //elements.add(e);
            //p.select("p[0]").remove();
            
            
            text [i]= strong_big[i] + ". " + ul.text();
            
            
            Elements day__text_em = ul.select("li em");
            ul.select("li em").remove();
            Elements day__text_strong = ul.select("li strong");
            Elements day__text_b = ul.select("li b");//new version: ("div.day__text b");
            Elements day__text_span = ul.select("li span");//> *:not(em)");//p.select("a span :not(a em)");
            
            //String[] strong_arr = new String[day__text_strong.size()];
            for (int j = 0; j < day__text_strong.size(); j++) {
                //strong_arr[j] = day__text_strong.get(j).text();
                strong_matr[i][j] = day__text_strong.get(j).text();
            }
            for (int j = 0; j < day__text_em.size(); j++) {
                //strong_arr[j] = day__text_strong.get(j).text();
                em_matr[i][j] = day__text_em.get(j).text();
            }
            for (int j = 0; j < day__text_b.size(); j++) {
                //strong_arr[j] = day__text_strong.get(j).text();
                strong_matr[i][j+day__text_strong.size()] = day__text_b.get(j).text();
            }
            for (int j = 0; j < day__text_span.size(); j++) {
                //strong_arr[j] = day__text_strong.get(j).text();
                if (day__text_span.get(j).text().length()!=0)//new version: day__text_span.get(j).text()!=0
                strong_matr[i][j+day__text_strong.size()] = day__text_span.get(j).text();
            }
            
            //int qwe = strong_matr[i].length;
            //System.out.println("Title : " + qwe);
            
            
            System.out.println("Title : " + text[i]);
            System.out.println("DAY_OF_MONTH: " + calendar.get(Calendar.DAY_OF_MONTH));
            date_full[i] = df1.format(calendar.getTime());
            System.out.println("DAY_OF_MONTH: " + date_full[i]);//calendar.get(Calendar.DAY_OF_MONTH));
            calendar.add(Calendar.DAY_OF_MONTH, +1);
        }
               
        //System.out.println(calendar.getTime());

        //Path templatePath = Paths.get(Timetable.class.getClassLoader().getResource(filename).toURI());
        //XWPFDocument doc1 =  new XWPFDocument(Files.newInputStream(templatePath));
        XWPFDocument doc1 = new XWPFDocument(OPCPackage.open("template.docx"));
        //XWPFTable tabl = doc1.createTable();
        String [] date_check = {"mon","tue","wen","thu","fri","sat","sun"}; //String date_check;
        String [] text_check = {"one","two","three","four","five","six","seven"}; //String text_check;
        for (i = 0; i < 7; i++) {
            //date_check = "month"+i;
           // text_check = i+"text";
           //String[] strong_arr = new strong[i][0];
            doc1 = replaceTextFor(doc1, date_check[i], date_full[i], null, null, i, strong_big[i]);
            doc1 = replaceTextFor(doc1, text_check[i], text[i], strong_matr, em_matr, i, strong_big[i]);     
        }
        doc1.write(new FileOutputStream("Расписание на " + Output_file + ".docx"));
        //try (XWPFDocument doc2 = new XWPFDocument(
            //Files.newInputStream(Paths.get(filename)))) {
            //XWPFWordExtractor extractor = new XWPFWordExtractor(doc2);    
            //String docText = extractor.getText();
            //System.out.println(docText);

            //// find number of words in the document
            ////long count = Arrays.stream(docText.split("\\s+")).count();
            ////System.out.println("Total words: " + count);

        //}
        
    }
        
    static void cloneRunProperties(XWPFRun source, XWPFRun dest) { // clones the underlying w:rPr element
        CTR tRSource = source.getCTR();
        CTRPr rPrSource = tRSource.getRPr();
        if (rPrSource != null) {
            CTRPr rPrDest = (CTRPr) rPrSource.copy();
            CTR tRDest = dest.getCTR();
            tRDest.setRPr(rPrDest);
        }
    }

    static void formatWord(XWPFParagraph paragraph, String keyword, Map<String, String> formats) {
        int runNumber = 0;
        while (runNumber < paragraph.getRuns().size()) { //go through all runs, we cannot use for each since we will possibly insert new runs
            XWPFRun run = paragraph.getRuns().get(runNumber);
            XWPFRun run2 = run;
            String runText = run.getText(0);
            if (runText != null && runText.contains(keyword)) { //if we have a run with keyword in it, then

                // This code part is to manage comment ranges.
                // Do we have commentRangeEnd immediately after the run?
                // If so then remember that in a cursor.
                /**
                 * XmlCursor commentRangeEndCursor = null; XmlCursor cursor =
                 * run.getCTR().newCursor(); cursor.toEndToken(); if
                 * (cursor.hasNextToken()) { cursor.toNextToken(); XmlObject
                 * commentRangeEnd = cursor.getObject(); if (commentRangeEnd !=
                 * null && commentRangeEnd instanceof CTMarkupRange) {
                 * commentRangeEndCursor = cursor; } }
                 */
                char[] runChars = runText.toCharArray(); //split run text into characters
                StringBuffer sb = new StringBuffer();
                for (int charNumber = 0; charNumber < runChars.length; charNumber++) { //go through all characters in that run
                    sb.append(runChars[charNumber]); //buffer all characters
                    runText = sb.toString();
                    if (runText.endsWith(keyword)) { //if the bufferend character stream ends with the keyword  
                        //set all chars, which are current buffered, except the keyword, as the text of the actual run
                        run.setText(runText.substring(0, runText.length() - keyword.length()), 0);
                        run2 = paragraph.insertNewRun(++runNumber); //insert new run for the formatted keyword
                        cloneRunProperties(run, run2); // clone the run properties from original run
                        run2.setText(keyword, 0); // set the keyword in run
                        for (String toSet : formats.keySet()) { // do the additional formatting
                            if ("italic".equals(toSet)) {
                                run2.setItalic(Boolean.valueOf(formats.get(toSet)));
                            } else if ("bold".equals(toSet)) {
                                run2.setBold(Boolean.valueOf(formats.get(toSet)));
                            } else if ("fontsize".equals(toSet)) {
                                run2.setFontSize(11);
                            }
                        }
                        run2 = paragraph.insertNewRun(++runNumber); //insert a new run for the next characters
                        cloneRunProperties(run, run2); // clone the run properties from original run
                        run = run2;
                        sb = new StringBuffer(); //empty the buffer
                    }
                }
                run.setText(sb.toString(), 0); //set all characters, which are currently buffered, as the text of the actual run

                // This code part is to manage comment ranges.
                // If we had remembered commentRangeEnd, then move this to here now.
                /**
                 * if(commentRangeEndCursor != null) { cursor =
                 * run.getCTR().newCursor(); cursor.toEndToken(); if
                 * (cursor.hasNextToken()) { cursor.toNextToken();
                 * commentRangeEndCursor.moveXml(cursor); } cursor.dispose();
                 * commentRangeEndCursor.dispose(); }
                 */
            }
            runNumber++;
        }
    }
    
    
    
    
    private static XWPFDocument replaceTextFor(XWPFDocument doc3, String findText, String replaceText, String [][] keyword_strong, String [][] keyword_em, int num, String strong_big) throws  IOException, InvalidFormatException 
    {
       
                //XWPFDocument doc3 = new XWPFDocument(OPCPackage.open("template.docx"));
                /*
       for (XWPFParagraph p1 : doc3.getParagraphs()) {
           List<XWPFRun> runs = p1.getRuns();
           if (runs != null) {
               for (XWPFRun r : runs) {
                   String text = r.getText(0);
                   if (text != null && text.contains(findText)){
                       text = text.replace(findText, replaceText);
                       r.setText(text, 0);
                   }
               }
           }
       }
       */

       Map<String, String> formats = new HashMap<String, String>();
       for (XWPFTable tbl : doc3.getTables()) {
          for (XWPFTableRow row : tbl.getRows()) {
             for (XWPFTableCell cell : row.getTableCells()) {
                for (XWPFParagraph p1 : cell.getParagraphs()) {
                   for (XWPFRun r : p1.getRuns()) {
                     String text = r.getText(0);
                     if (text != null && text.contains(findText)) {
                       text = text.replace(findText, replaceText);
                       r.setText(text,0);
                     }
                   }
                   if (keyword_strong!=null)
                   if (keyword_strong[num][0]!=null)
                   //if (keyword_strong[num][0].length()!=0)
                   {
                       int i = 0;
                       while (keyword_strong[num][i]!=null)
                       {
                           formats.put("bold", "true");
                           formats.put("italic", "false");
                           //formats.put("fontsize", "false");
                           formatWord(p1, keyword_strong[num][i], formats);
                           i++;
                       }


                   } 
                   if (keyword_em!=null)
                   if (keyword_em[num][0]!=null)
                   if (keyword_em[num][0].length()!=0)
                   {
                       int i = 0;
                       while (keyword_em[num][i]!=null)
                       {
                           formats.put("bold", "false");
                           formats.put("italic", "true");
                           //formats.put("fontsize", "false");
                           formatWord(p1, keyword_em[num][i], formats);
                           i++;
                       } 
                   }  
                   if (strong_big!=null)
                   if (strong_big.length()!=0)
                   {
                           formats.put("bold", "true");
                           formats.put("italic", "false");
                           formats.put("fontsize", "true");
                           formatWord(p1, strong_big, formats);
                           formats = new HashMap<String, String>();
                   }

                }
             }
          }
       }
              return doc3;
        
        
    }

    private static XWPFDocument replaceTable(XWPFDocument doc3, XWPFTable table, String placeHolder, String replaceText) {
    for (XWPFTableRow row : table.getRows()) {
        for (XWPFTableCell cell : row.getTableCells()) {
            for (IBodyElement bodyElement : cell.getBodyElements()) {
                if (bodyElement.getElementType().compareTo(BodyElementType.PARAGRAPH) == 0) {
                    replaceParagraph((XWPFParagraph) bodyElement, placeHolder, replaceText);
                }
                if (bodyElement.getElementType().compareTo(BodyElementType.TABLE) == 0) {
                    replaceTable(doc3,(XWPFTable) bodyElement, placeHolder, replaceText);
                }
            }
        }
    }
    return doc3;
}
    private static void replaceParagraph(XWPFParagraph paragraph, String placeHolder, String replaceText) {
    for (XWPFRun r : paragraph.getRuns()) {
        String text = r.getText(r.getTextPosition());
        if (text != null && text.contains(placeHolder)) {
            text = text.replace(placeHolder, replaceText);
            r.setText(text, 0);
        }
    }
    }
    
    
    private static DateFormatSymbols myDateFormatSymbols = new DateFormatSymbols(){

        @Override
        public String[] getMonths() {
            return new String[]{"января", "февраля", "марта", "апреля", "мая", "июня",
                "июля", "августа", "сентября", "октября", "ноября", "декабря"};
        }
        
    };

    private static void add(Element ptemp) {
        throw new UnsupportedOperationException("Not supported yet."); //To change body of generated methods, choose Tools | Templates.
    }
   
    
}
