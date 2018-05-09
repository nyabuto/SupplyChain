/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Reports;

import Db.dbConn;
import SupplyChain.Manager;
import java.io.ByteArrayOutputStream;
import java.io.IOException;
import java.io.OutputStream;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import org.apache.poi.hssf.util.HSSFColor;
import org.apache.poi.ss.usermodel.BorderStyle;
import org.apache.poi.ss.usermodel.CellStyle;
import org.apache.poi.ss.usermodel.DataFormat;
import org.apache.poi.ss.usermodel.FillPatternType;
import org.apache.poi.ss.usermodel.FontFamily;
import org.apache.poi.ss.usermodel.HorizontalAlignment;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFCellStyle;
import org.apache.poi.xssf.usermodel.XSSFFont;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

/**
 *
 * @author GNyabuto
 */
public class raw_report extends HttpServlet {
HttpSession session;
String query;
int row_counter;
String where_clause,columns;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
        
            session = request.getSession();
            dbConn conn = new dbConn();
          Manager manager = new Manager();
           
          String[] counties = request.getParameterValues("county");
          String[] sub_counties = request.getParameterValues("sub_county");
          String[] facilities = request.getParameterValues("mfl_code");
          
          String[] elements = request.getParameterValues("elements");
          
          columns="";
          
         where_clause = "WHERE 1=1 AND"; 
          if(facilities!=null){
          for(String mfl_code:facilities){
              if(!mfl_code.equals("") && !mfl_code.equals(",")){
              mfl_code = mfl_code.replace("'", "\'");
              where_clause+=" mfl_code='"+mfl_code+"' OR ";
              }
          }
          }
          
          else if(sub_counties!=null){
          for(String sub_county:sub_counties){
              if(!sub_county.equals("") && !sub_county.equals(",")){
              sub_county = sub_county.replace("'", "\'");
              where_clause+=" sub_county='"+sub_county+"' OR ";
              }
          }
          }
          
          else if(counties!=null){
          for(String county:counties){
              if(!county.equals("") && !county.equals(",")){
              county = county.replace("'", "\'");
              where_clause+=" county='"+county+"' OR ";
              }
          }
          }
          
          if(elements!=null){
          for(String column_name:elements){
              if(!column_name.equals("") && !column_name.equals(",")){
              column_name = column_name.replace("'", "\'");
              String label = getcolum_label(conn,column_name);
               label = label.replace("'", "\\'");
              columns+="IFNULL("+column_name+",'') AS '"+label+"', ";
              }
          }
          }
          else{
           String getlabel  = "SELECT column_name,label FROM column_mapping WHERE is_active=1 ORDER BY id";
       conn.rs = conn.st.executeQuery(getlabel);
       while(conn.rs.next()){
           String column_name = conn.rs.getString(1);
           String label = conn.rs.getString(2);
           
           column_name = column_name.replace("'", "\'");
           label = label.replace("'", "\\'");
           columns+="IFNULL("+column_name+",'') AS '"+label+"', ";
       }   
          }
          
          
         columns = Manager.removeLastChars(columns, 2); 
          
          //remove last 4 characters
          where_clause = Manager.removeLastChars(where_clause, 3);
            row_counter=0;
      
        //            ^^^^^^^^^^^^^CREATE STATIC AND WRITE STATIC DATA TO THE EXCELL^^^^^^^^^^^^
    XSSFWorkbook wb=new XSSFWorkbook();
    XSSFSheet shet1=wb.createSheet("Baseline Assessment Data");
    XSSFFont font=wb.createFont();
    font.setFontHeightInPoints((short)18);
    font.setFontName("Cambria");
    font.setColor((short)0000);
    
    CellStyle style=wb.createCellStyle();
    style.setFont(font);
    style.setAlignment(HorizontalAlignment.CENTER);
    
    XSSFCellStyle styleHeader = wb.createCellStyle();
    styleHeader.setFillForegroundColor(HSSFColor.PALE_BLUE.index);
    styleHeader.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    styleHeader.setBorderTop(BorderStyle.THIN);
    styleHeader.setBorderBottom(BorderStyle.THIN);
    styleHeader.setBorderLeft(BorderStyle.THIN);
    styleHeader.setBorderRight(BorderStyle.THIN);
    styleHeader.setAlignment(HorizontalAlignment.CENTER);
    
    XSSFFont fontHeader = wb.createFont();
    fontHeader.setColor(HSSFColor.BLACK.index);
    fontHeader.setBold(true);
    fontHeader.setFamily(FontFamily.MODERN);
    fontHeader.setFontName("Cambria");
    fontHeader.setFontHeight(12);
    styleHeader.setFont(fontHeader);
    styleHeader.setWrapText(true);
    
    XSSFCellStyle stylex = wb.createCellStyle();
    stylex.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
    stylex.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    stylex.setBorderTop(BorderStyle.THIN);
    stylex.setBorderBottom(BorderStyle.THIN);
    stylex.setBorderLeft(BorderStyle.THIN);
    stylex.setBorderRight(BorderStyle.THIN);
    stylex.setAlignment(HorizontalAlignment.LEFT);
    
    XSSFFont fontx = wb.createFont();
    fontx.setColor(HSSFColor.BLACK.index);
    fontx.setBold(true);
    fontx.setFamily(FontFamily.MODERN);
    fontx.setFontName("Cambria");
    stylex.setFont(fontx);
    stylex.setWrapText(true);
    
    XSSFCellStyle stborder = wb.createCellStyle();
    stborder.setBorderTop(BorderStyle.THIN);
    stborder.setBorderBottom(BorderStyle.THIN);
    stborder.setBorderLeft(BorderStyle.THIN);
    stborder.setBorderRight(BorderStyle.THIN);
    stborder.setAlignment(HorizontalAlignment.LEFT);
    
    XSSFFont font_cell=wb.createFont();
    font_cell.setColor(HSSFColor.BLACK.index);
    font_cell.setFamily(FontFamily.MODERN);
    font_cell.setFontName("Cambria");
    stborder.setFont(font_cell);
    stborder.setWrapText(true);
    
    XSSFCellStyle currencyStyle=wb.createCellStyle();
    currencyStyle.setBorderTop(BorderStyle.THIN);
    currencyStyle.setBorderBottom(BorderStyle.THIN);
    currencyStyle.setBorderLeft(BorderStyle.THIN);
    currencyStyle.setBorderRight(BorderStyle.THIN);
    currencyStyle.setAlignment(HorizontalAlignment.LEFT);
    
    currencyStyle.setFont(font_cell);
    currencyStyle.setWrapText(true);
    
    XSSFCellStyle totalcurrencyStyle=wb.createCellStyle();
    totalcurrencyStyle.setFillForegroundColor(HSSFColor.GREY_25_PERCENT.index);
    totalcurrencyStyle.setFillPattern(FillPatternType.SOLID_FOREGROUND);
    totalcurrencyStyle.setBorderTop(BorderStyle.THIN);
    totalcurrencyStyle.setBorderBottom(BorderStyle.THIN);
    totalcurrencyStyle.setBorderLeft(BorderStyle.THIN);
    totalcurrencyStyle.setBorderRight(BorderStyle.THIN);
    totalcurrencyStyle.setAlignment(HorizontalAlignment.LEFT);
    
    totalcurrencyStyle.setFont(fontx);
    totalcurrencyStyle.setWrapText(true);
    
    
   DataFormat df = wb.createDataFormat();
   
   // fetch currency
          
            
            
            
            
            
        query = " SELECT "
                +columns+
                " FROM report  "+where_clause+"";
        
        conn.pst = conn.conn.prepareStatement(query);
        conn.rs = conn.pst.executeQuery();
        
        
        
        ResultSetMetaData metaData = conn.rs.getMetaData();
        int count = metaData.getColumnCount(); //number of column
        
        XSSFRow rw0=shet1.createRow(row_counter); 
        rw0.setHeightInPoints(35);
        int j=1;
           while(j<=count){
               XSSFCell  S1cell=rw0.createCell(j-1);
               S1cell.setCellValue(metaData.getColumnLabel(j));
               S1cell.setCellStyle(styleHeader);
//                   System.out.println("columns : "+metaData.getColumnLabel(j));   
               j++;
           }
          
 for (int i=0;i<count;i++){
   shet1.setColumnWidth(i, 6000);
    }
               
               
        while(conn.rs.next()){
            row_counter++;
          XSSFRow rw=shet1.createRow(row_counter);   
            
          int i=1;
           while(i<=count){
               String value = conn.rs.getString(i);
               XSSFCell  cell=rw.createCell(i-1);
               if(isNumeric(value)){
                cell.setCellValue(Double.parseDouble(value));
            }
            else{
                cell.setCellValue(value);
            }
               cell.setCellStyle(stborder);
//               System.out.println(metaData.getColumnLabel(i)+"Value : "+conn.rs.getString(i));                 
               i++;
           }
        }
        
        ByteArrayOutputStream outByteStream = new ByteArrayOutputStream();
        wb.write(outByteStream);
        byte [] outArray = outByteStream.toByteArray();
        response.setContentType("application/ms-excel");
        response.setContentLength(outArray.length);
        response.setHeader("Expires:", "0"); // eliminates browser caching
        response.setHeader("Content-Disposition", "attachment; filename=SupplyChain_Checklist_Raw_Data_"+manager.getdatekey()+".xlsx");
        OutputStream outStream = response.getOutputStream();
        outStream.write(outArray);
        outStream.flush();          
    }

    // <editor-fold defaultstate="collapsed" desc="HttpServlet methods. Click on the + sign on the left to edit the code.">
    /**
     * Handles the HTTP <code>GET</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doGet(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
    try {
        processRequest(request, response);
    } catch (SQLException ex) {
        Logger.getLogger(raw_report.class.getName()).log(Level.SEVERE, null, ex);
    }
    }

    /**
     * Handles the HTTP <code>POST</code> method.
     *
     * @param request servlet request
     * @param response servlet response
     * @throws ServletException if a servlet-specific error occurs
     * @throws IOException if an I/O error occurs
     */
    @Override
    protected void doPost(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException {
    try {
        processRequest(request, response);
    } catch (SQLException ex) {
        Logger.getLogger(raw_report.class.getName()).log(Level.SEVERE, null, ex);
    }
    }

    /**
     * Returns a short description of the servlet.
     *
     * @return a String containing servlet description
     */
    @Override
    public String getServletInfo() {
        return "Short description";
    }// </editor-fold>

    public boolean isNumeric(String s) {  
        return s != null && s.matches("[-+]?\\d*\\.?\\d+");  
}
  
    public String getcolum_label(dbConn conn, String column_name) throws SQLException{
       String label = "";
       String getlabel  = "SELECT label FROM column_mapping WHERE column_name='"+column_name+"'";
       conn.rs = conn.st.executeQuery(getlabel);
       if(conn.rs.next()){
           label = conn.rs.getString(1);
       }
       
       return label;
    }
}
