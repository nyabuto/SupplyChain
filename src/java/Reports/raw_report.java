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
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
        
            session = request.getSession();
            dbConn conn = new dbConn();
          Manager manager = new Manager();
            
            row_counter=0;
      
        //            ^^^^^^^^^^^^^CREATE STATIC AND WRITE STATIC DATA TO THE EXCELL^^^^^^^^^^^^
    XSSFWorkbook wb=new XSSFWorkbook();
    XSSFSheet shet1=wb.createSheet("Rebanking Report");
    XSSFFont font=wb.createFont();
    font.setFontHeightInPoints((short)18);
    font.setFontName("Cambria");
    font.setColor((short)0000);
    
    CellStyle style=wb.createCellStyle();
    style.setFont(font);
    style.setAlignment(HorizontalAlignment.CENTER);
    
    XSSFCellStyle styleHeader = wb.createCellStyle();
    styleHeader.setFillForegroundColor(HSSFColor.GREY_40_PERCENT.index);
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
    fontHeader.setFontHeight(15);
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
          
            
            
            
            
            
        query = " SELECT IFNULL(county,'') AS 'County', IFNULL(sub_county,'') AS 'Sub County',IFNULL(facility,'') AS'Health Facility', "
            + "IFNULL(facility_level,'') AS 'Facility Level',  IFNULL(mfl_code,'') AS 'MFL Code', IFNULL(ownership,'') AS 'Ownership', "
            + "IFNULL(facility_incharge,'') AS 'Facility in Charge', IFNULL(incharge_contact,'') AS 'Facility in Charge Phone', "
            + "IFNULL(pharmacy_person,'') AS 'Pharmacy Contact Person', IFNULL(pharmacy_phone,'') AS 'Pharmacy contact person\\'s Phone', "
            + "IFNULL(lab_person,'') AS 'Laboratory Contact Person', IFNULL(lab_phone,'') AS 'Laboratory contact Person\\'s Phone', "
            + "IFNULL(visit_date,'') AS 'Date of Visit', IFNULL(has_lab,'') AS 'Does the facility have a laboratory?', "
            + "IFNULL(has_designated_store,'') AS 'Does the facility have a designated store for health commodities?', "
            + "IFNULL(CM_pharmaceutical,'') AS 'Who manages Pharmaceuticals and related supplies?', "
            + "IFNULL(CM_non_pharmaceutical,'') AS 'Who manages Non-pharmaceuticals (medical supplies)?', "
            + "IFNULL(CM_lab,'') AS 'Who manages Lab reagents and test kits?', "
            + "IFNULL(supervision_CM,'') AS 'Has the facility received supportive supervision related to commodity management in the last 3 months?', "
            + "IFNULL(who_supervision_CM,'') AS 'If yes above who provided this supportive supervision?', "
            + "IFNULL(power_source,'') AS 'What power source does the facility use (Electricity, Solar, etc.)?', "
            + "IFNULL(power_reliable,'') AS 'Is the power supply reliable (not more than one outage of >5 minutes per day)?', "
            + "IFNULL(commodity_challenges,'') AS 'What are the main commodity related challenges that the facility experiences?', "
            + "IFNULL(storage_pharm_num,'') AS 'Pharmacy storage area numerator', IFNULL(storage_pharm_den,'') AS 'Pharmacy storage area denominator', "
            + "IFNULL(storage_pharm_score,'') AS 'Pharmacy storage area score', IFNULL(storage_lab_num,'') AS 'Laboratory starage area numerator', "
            + "IFNULL(storage_lab_den,'') AS 'Laboratory storage area denominator', IFNULL(storage_lab_score,'') AS 'Laboratory storage area score', "
            + "IFNULL(storage_total_num,'') AS 'Storage areas numerator', IFNULL(storage_total_den,'') AS 'Storage areas denominator', "
            + "IFNULL(storage_total_score,'') AS 'Storage areas Score', IFNULL(inventory_pharm_num,'') AS 'Pharmacy inventory management numerator', "
            + "IFNULL(inventory_pharm_den,'') AS 'Pharmacy inventory management denominator', IFNULL(inventory_pharm_score,'') AS 'Pharmacy inventory management score', "
            + "IFNULL(inventory_lab_num,'') AS 'Laboratory inventory management numerator', IFNULL(inventory_lab_den,'') AS 'Laboratory inventory management denominator', "
            + "IFNULL(inventory_lab_score,'') AS 'Laboratory inventory management score', IFNULL(inventory_total_num,'') AS 'Inventory management numerator', "
            + "IFNULL(inventory_total_den,'') AS 'Inventory management denominator', IFNULL(inventory_total_score,'') AS 'Inventory management score', "
            + "IFNULL(RRM_pharm_num,'') AS 'Pharmacy resources & reference materials numerator', "
            + "IFNULL(RRM_pharm_den,'') AS 'Pharmacy resources & reference materials denominator', IFNULL(RRM_pharm_score,'') AS 'Pharmacy resources & reference materials score', "
            + "IFNULL(RRM_lab_num,'') AS 'Laboratory resources & reference materials numerator', IFNULL(RRM_lab_den,'') AS 'Laboratory resources & reference materials denominator', "
            + "IFNULL(RRM_lab_score,'') AS 'Laboratory resources & reference materials score', IFNULL(RRM_total_num,'') AS 'Resources & reference materials numerator',"
            + " IFNULL(RRM_total_den,'') AS 'Resources & reference materials denominator', IFNULL(RRM_total_score,'') AS 'Resources & reference materials score', "
            + "IFNULL(MIS_pharm_num,'') AS 'Availability & Use of MIS Tools in pharmacy numerator', "
            + "IFNULL(MIS_pharm_den,'') AS 'Availability & Use of MIS Tools in pharmacy denominator', "
            + "IFNULL(MIS_pharm_score,'') AS 'Availability & Use of MIS Tools in pharmacy score', IFNULL(MIS_lab_num,'') AS 'Availability & Use of MIS Tools in laboratory numerator', "
            + "IFNULL(MIS_lab_den,'') AS 'Availability & Use of MIS Tools in laboratory denominator', IFNULL(MIS_lab_score,'') AS 'Availability & Use of MIS Tools in laboratory score', "
            + "IFNULL(MIS_total_num,'') AS 'Availability & Use of MIS Tools numerator', IFNULL(MIS_total_den,'') AS 'Availability & Use of MIS Tools denominator', "
            + "IFNULL(MIS_total_score,'') AS 'Availability & Use of MIS Tools score', IFNULL(total_pharm_num,'') AS 'Pharmacy total numerator', "
            + "IFNULL(total_pharm_den,'') AS 'Pharmacy total denominator', IFNULL(total_pharm_score,'') AS 'Pharmacy total score', "
            + "IFNULL(total_lab_num,'') AS 'Laboratory total numerator', IFNULL(total_lab_den,'') AS 'Laboratory total denominator', "
            + "IFNULL(total_lab_score,'') AS 'Laboratory total score', IFNULL(total_num,'') AS 'Total Numerator', IFNULL(total_den,'') AS 'Total denominator', "
            + "IFNULL(total_score,'') AS 'Total score', IFNULL(EUV_pharm_num,'') AS 'Pharmacy HIV RTKs End Use Verification numerator', "
            + "IFNULL(EUV_pharm_den,'') AS 'Pharmacy HIV RTKs End Use Verification denominator', IFNULL(EUV_pharm_score,'') AS 'Pharmacy HIV RTKs End Use Verification score ', "
            + "IFNULL(EUV_lab_num,'') AS 'Laboratory HIV RTKs End Use Verification numerator', IFNULL(EUV_lab_den,'') AS 'Laboratory HIV RTKs End Use Verification denominator', "
            + "IFNULL(EUV_lab_score,'') AS 'Laboratory HIV RTKs End Use Verification score ', "
            + "IFNULL(IM_additional_pharm_num,'') AS 'Pharmacy Inventory Management: additional commodities numerator', "
            + "IFNULL(IM_additional_pharm_den,'') AS 'Pharmacy Inventory Management: additional commodities denominator', "
            + "IFNULL(IM_additional_pharm_score,'') AS 'Pharmacy Inventory Management: additional commodities score', "
            + "IFNULL(IM_additional_lab_num,'') AS 'Laboratory Inventory Management: additional commodities numerator', "
            + "IFNULL(IM_additional_lab_den,'') AS 'Laboratory Inventory Management: additional commodities denominator', "
            + "IFNULL(IM_additional_lab_score,'') AS 'Laboratory Inventory Management: additional commodities score', "
            + "IFNULL(MTC_pharm_num,'') AS 'Pharmacy Medicines and Therapeutics Committees numerator', "
            + "IFNULL(MTC_pharm_den,'') AS 'Pharmacy Medicines and Therapeutics Committees denominator', "
            + "IFNULL(MTC_pharm_score,'') AS 'Pharmacy Medicines and Therapeutics Committees Score', "
            + "IFNULL(MTC_lab_num,'') AS 'Laboratory Medicines and Therapeutics Committees numerator', "
            + "IFNULL(MTC_lab_den,'') AS 'Laboratory Medicines and Therapeutics Committees denominator ', "
            + "IFNULL(MTC_lab_score,'') AS 'Laboratory Medicines and Therapeutics Committees score' FROM report ";
        
        conn.pst = conn.conn.prepareStatement(query);
        conn.rs = conn.pst.executeQuery();
        
        
        
        ResultSetMetaData metaData = conn.rs.getMetaData();
        int count = metaData.getColumnCount(); //number of column
        
        XSSFRow rw0=shet1.createRow(row_counter); 
        
        int j=1;
           while(j<=count){
               XSSFCell  S1cell=rw0.createCell(j-1);
               S1cell.setCellValue(metaData.getColumnLabel(j));
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
        response.setHeader("Content-Disposition", "attachment; filename=Rebanking_Report_"+manager.getdatekey()+".xlsx");
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
    
}
