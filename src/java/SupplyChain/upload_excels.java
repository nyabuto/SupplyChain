/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package SupplyChain;

import Db.dbConn;
import java.io.File;
import java.io.FileInputStream;
import java.io.IOException;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.servlet.ServletException;
import javax.servlet.annotation.MultipartConfig;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import javax.servlet.http.Part;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

/**
 *
 * @author GNyabuto
 */
@MultipartConfig
public class upload_excels extends HttpServlet {
  String full_path="";
  String fileName="";
  File file_source;
  HttpSession session;
  private static final long serialVersionUID = 205242440643911308L;
  private static final String UPLOAD_DIR = "uploads";
  String query="",value,checker_query;
  String id,facility,facility_level,county,sub_county,mfl_code,level,ownership,facility_incharge,incharge_contact,pharmacy_person,pharmacy_phone,lab_person,lab_phone,visit_date,has_lab,has_designated_store,CM_pharmaceutical,CM_non_pharmaceutical,CM_lab,supervision_CM,who_supervision_CM,power_source,power_reliable,commodity_challenges,storage_pharm_num,storage_pharm_den,storage_pharm_score,storage_lab_num,storage_lab_den,storage_lab_score,storage_total_num,storage_total_den,storage_total_score,inventory_pharm_num,inventory_pharm_den,inventory_pharm_score,inventory_lab_num,inventory_lab_den,inventory_lab_score,inventory_total_num,inventory_total_den,inventory_total_score,RRM_pharm_num,RRM_pharm_den,RRM_pharm_score,RRM_lab_num,RRM_lab_den,RRM_lab_score,RRM_total_num,RRM_total_den,RRM_total_score,MIS_pharm_num,MIS_pharm_den,MIS_pharm_score,MIS_lab_num,MIS_lab_den,MIS_lab_score,MIS_total_num,MIS_total_den,MIS_total_score,total_pharm_num,total_pharm_den,total_pharm_score,total_lab_num,total_lab_den,total_lab_score,total_num,total_den,total_score,EUV_pharm_num,EUV_pharm_den,EUV_pharm_score,EUV_lab_num,EUV_lab_den,EUV_lab_score,IM_additional_pharm_num,IM_additional_pharm_den,IM_additional_pharm_score,IM_additional_lab_num,IM_additional_lab_den,IM_additional_lab_score,MTC_pharm_num,MTC_pharm_den,MTC_pharm_score,MTC_lab_num,MTC_lab_den,MTC_lab_score,is_locked,timestamp;
int number_changes,added;
String failed_description;
int failed;
  protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
            dbConn  conn = new dbConn();
            session = request.getSession();
            JSONObject finalobj = new JSONObject();
            JSONArray jarray = new JSONArray();
            number_changes=added=0;
        
        String applicationPath = request.getServletContext().getRealPath("");
         String uploadFilePath = applicationPath + File.separator + UPLOAD_DIR;
         session=request.getSession();
          File fileSaveDir = new File(uploadFilePath);
        if (!fileSaveDir.exists()) {
            fileSaveDir.mkdirs();
        }
        
        for (Part part : request.getParts()) {
            
            if(!getFileName(part).equals("")){
           fileName = getFileName(part);
            part.write(uploadFilePath + File.separator + fileName);
            
            full_path=fileSaveDir.getAbsolutePath()+"\\"+fileName;
            System.out.println("fullpath : "+full_path);
           // read the contents of the workbook and sheets here
           
           FileInputStream fileInputStream = new FileInputStream(full_path);
        XSSFWorkbook workbook = new XSSFWorkbook(fileInputStream);
        int j=0;
//        int number_sheets = workbook.getNumberOfSheets();
        XSSFSheet worksheetSummary,workSheetGeneral;
        
        workSheetGeneral = workbook.getSheet("General");
        worksheetSummary = workbook.getSheet("Summary Scores");
//        worksheetSummary = workbook.getSheetAt(6);
                
        failed=0;
        failed_description="";
                
//****************************FACILITY INFORMATION****************************************
//                facility
                if(null==workSheetGeneral.getRow(1).getCell(4).getCellTypeEnum()){
                    facility = workSheetGeneral.getRow(1).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(1).getCell(4).getCellTypeEnum()) {
                    case NUMERIC:
                        facility = ""+workSheetGeneral.getRow(1).getCell(4).getNumericCellValue();
                        break;
                    case STRING:
                        facility = workSheetGeneral.getRow(1).getCell(4).getStringCellValue();
                        break;
                    case FORMULA:
                        facility = ""+workSheetGeneral.getRow(1).getCell(4).getNumericCellValue();
                        break;
                    default:
                        facility = workSheetGeneral.getRow(1).getCell(4).getRawValue();
                        break;
                }
//        end facility

//                Sub County
                if(null==workSheetGeneral.getRow(1).getCell(12).getCellTypeEnum()){
                    sub_county = workSheetGeneral.getRow(1).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(1).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        sub_county = ""+workSheetGeneral.getRow(1).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        sub_county = workSheetGeneral.getRow(1).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        sub_county = ""+workSheetGeneral.getRow(1).getCell(12).getNumericCellValue();
                        break;
                    default:
                        sub_county = workSheetGeneral.getRow(1).getCell(12).getRawValue();
                        break;
                }
//        end sub county

//                County
                if(null==workSheetGeneral.getRow(2).getCell(4).getCellTypeEnum()){
                    county = workSheetGeneral.getRow(2).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(2).getCell(4).getCellTypeEnum()) {
                    case NUMERIC:
                        county = ""+workSheetGeneral.getRow(2).getCell(4).getNumericCellValue();
                        break;
                    case STRING:
                        county = workSheetGeneral.getRow(2).getCell(4).getStringCellValue();
                        break;
                    case FORMULA:
                        county = ""+workSheetGeneral.getRow(2).getCell(4).getNumericCellValue();
                        break;
                    default:
                        county = workSheetGeneral.getRow(2).getCell(4).getRawValue();
                        break;
                }
//        end county

//                 facility_level
                if(null==workSheetGeneral.getRow(2).getCell(12).getCellTypeEnum()){
                    facility_level = workSheetGeneral.getRow(2).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(2).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        facility_level = ""+workSheetGeneral.getRow(2).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        facility_level = workSheetGeneral.getRow(2).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        facility_level = ""+workSheetGeneral.getRow(2).getCell(12).getNumericCellValue();
                        break;
                    default:
                        facility_level = workSheetGeneral.getRow(2).getCell(12).getRawValue();
                        break;
                }
//        end facility_level

//                mfl_code
                if(null==workSheetGeneral.getRow(3).getCell(4).getCellTypeEnum()){
                    mfl_code = workSheetGeneral.getRow(3).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(3).getCell(4).getCellTypeEnum()) {
                    case NUMERIC:
                        mfl_code = ""+workSheetGeneral.getRow(3).getCell(4).getNumericCellValue();
                        break;
                    case STRING:
                        mfl_code = workSheetGeneral.getRow(3).getCell(4).getStringCellValue();
                        break;
                    case FORMULA:
                        mfl_code = ""+workSheetGeneral.getRow(3).getCell(4).getNumericCellValue();
                        break;
                    default:
                        mfl_code = workSheetGeneral.getRow(3).getCell(4).getRawValue();
                        break;
                }
//        end mfl_code

//                ownership
                if(null==workSheetGeneral.getRow(4).getCell(4).getCellTypeEnum()){
                    ownership = workSheetGeneral.getRow(4).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(4).getCell(4).getCellTypeEnum()) {
                    case NUMERIC:
                        ownership = ""+workSheetGeneral.getRow(4).getCell(4).getNumericCellValue();
                        break;
                    case STRING:
                        ownership = workSheetGeneral.getRow(4).getCell(4).getStringCellValue();
                        break;
                    case FORMULA:
                        ownership = ""+workSheetGeneral.getRow(4).getCell(4).getNumericCellValue();
                        break;
                    default:
                        ownership = workSheetGeneral.getRow(4).getCell(4).getRawValue();
                        break;
                }
//        end ownership


//                facility_incharge
                if(null==workSheetGeneral.getRow(5).getCell(4).getCellTypeEnum()){
                    facility_incharge = workSheetGeneral.getRow(5).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(5).getCell(4).getCellTypeEnum()) {
                    case NUMERIC:
                        facility_incharge = ""+workSheetGeneral.getRow(5).getCell(4).getNumericCellValue();
                        break;
                    case STRING:
                        facility_incharge = workSheetGeneral.getRow(5).getCell(4).getStringCellValue();
                        break;
                    case FORMULA:
                        facility_incharge = ""+workSheetGeneral.getRow(5).getCell(4).getNumericCellValue();
                        break;
                    default:
                        facility_incharge = workSheetGeneral.getRow(5).getCell(4).getRawValue();
                        break;
                }
//        end facility_incharge

//                 incharge_contact
                    incharge_contact = workSheetGeneral.getRow(5).getCell(12).getRawValue();
                
//        end incharge_contact


//                pharmacy_person
                if(null==workSheetGeneral.getRow(6).getCell(4).getCellTypeEnum()){
                    pharmacy_person = workSheetGeneral.getRow(6).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(6).getCell(4).getCellTypeEnum()) {
                    case NUMERIC:
                        pharmacy_person = ""+workSheetGeneral.getRow(6).getCell(4).getNumericCellValue();
                        break;
                    case STRING:
                        pharmacy_person = workSheetGeneral.getRow(6).getCell(4).getStringCellValue();
                        break;
                    case FORMULA:
                        pharmacy_person = ""+workSheetGeneral.getRow(6).getCell(4).getNumericCellValue();
                        break;
                    default:
                        pharmacy_person = workSheetGeneral.getRow(6).getCell(4).getRawValue();
                        break;
                }
//        end pharmacy_person

//                 pharmacy_phone
                    pharmacy_phone = workSheetGeneral.getRow(6).getCell(12).getRawValue();
                
//        end pharmacy_phone


//                lab_person
                if(null==workSheetGeneral.getRow(7).getCell(4).getCellTypeEnum()){
                    lab_person = workSheetGeneral.getRow(7).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(7).getCell(4).getCellTypeEnum()) {
                    case NUMERIC:
                        lab_person = ""+workSheetGeneral.getRow(7).getCell(4).getNumericCellValue();
                        break;
                    case STRING:
                        lab_person = workSheetGeneral.getRow(7).getCell(4).getStringCellValue();
                        break;
                    case FORMULA:
                        lab_person = ""+workSheetGeneral.getRow(7).getCell(4).getNumericCellValue();
                        break;
                    default:
                        lab_person = workSheetGeneral.getRow(7).getCell(4).getRawValue();
                        break;
                }
//        end lab_person

//                 lab_phone
                    lab_phone = workSheetGeneral.getRow(7).getCell(12).getRawValue();

//        end lab_phone
            

//                visit_date
                if(null==workSheetGeneral.getRow(8).getCell(4).getCellTypeEnum()){
                    visit_date = workSheetGeneral.getRow(8).getCell(4).getRawValue();
                }
                else switch (workSheetGeneral.getRow(8).getCell(4).getCellTypeEnum()) {
                    
                    case STRING:
                        visit_date = workSheetGeneral.getRow(8).getCell(4).getStringCellValue();
                        break;
                    case NUMERIC:
                        visit_date = ""+workSheetGeneral.getRow(8).getCell(4).getNumericCellValue();
                        break;
                    case FORMULA:
                        visit_date = ""+workSheetGeneral.getRow(8).getCell(4).getNumericCellValue();
                        break;
                    default:
                        visit_date = workSheetGeneral.getRow(8).getCell(4).getRawValue();
                        break;
                }
//        end lab_person
//******************************************************************************

//    ****************INTERVIEW WITH FACILITY IN CHARGE**********************


//                 Does the facility have a laboratory?										
                if(null==workSheetGeneral.getRow(13).getCell(12).getCellTypeEnum()){
                    has_lab = workSheetGeneral.getRow(13).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(13).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        has_lab = ""+workSheetGeneral.getRow(13).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        has_lab = workSheetGeneral.getRow(13).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        has_lab = ""+workSheetGeneral.getRow(13).getCell(12).getNumericCellValue();
                        break;
                    default:
                        has_lab = workSheetGeneral.getRow(13).getCell(12).getRawValue();
                        break;
                }
                if(has_lab!=null && has_lab.equalsIgnoreCase("Yes")){
                    has_lab="1";
                }
                else{
                   has_lab="0";  
                }
//        end Does the facility have a laboratory?										


//                 Does the facility have a designated store for health commodities?																				
                if(null==workSheetGeneral.getRow(14).getCell(12).getCellTypeEnum()){
                    has_designated_store = workSheetGeneral.getRow(14).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(14).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        has_designated_store = ""+workSheetGeneral.getRow(14).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        has_designated_store = workSheetGeneral.getRow(14).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        has_designated_store = ""+workSheetGeneral.getRow(14).getCell(12).getNumericCellValue();
                        break;
                    default:
                        has_designated_store = workSheetGeneral.getRow(14).getCell(12).getRawValue();
                        break;
                }
                if(has_designated_store!=null && has_designated_store.equalsIgnoreCase("Yes")){
                    has_designated_store="1";
                }
                else{
                   has_designated_store="0";  
                }
//        end Does the facility have a designated store for health commodities?																				

//                  Pharmaceuticals and related supplies?																				
                if(null==workSheetGeneral.getRow(16).getCell(12).getCellTypeEnum()){
                    CM_pharmaceutical = workSheetGeneral.getRow(16).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(16).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        CM_pharmaceutical = ""+workSheetGeneral.getRow(16).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        CM_pharmaceutical = workSheetGeneral.getRow(16).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        CM_pharmaceutical = ""+workSheetGeneral.getRow(16).getCell(12).getNumericCellValue();
                        break;
                    default:
                        CM_pharmaceutical = workSheetGeneral.getRow(16).getCell(12).getRawValue();
                        break;
                }
//        end  Pharmaceuticals and related supplies?																			

//                  Non-pharmaceuticals (medical supplies)?																				
                if(null==workSheetGeneral.getRow(17).getCell(12).getCellTypeEnum()){
                    CM_non_pharmaceutical = workSheetGeneral.getRow(17).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(17).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        CM_non_pharmaceutical = ""+workSheetGeneral.getRow(17).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        CM_non_pharmaceutical = workSheetGeneral.getRow(17).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        CM_non_pharmaceutical = ""+workSheetGeneral.getRow(17).getCell(12).getNumericCellValue();
                        break;
                    default:
                        CM_non_pharmaceutical = workSheetGeneral.getRow(17).getCell(12).getRawValue();
                        break;
                }
//        end  Non-pharmaceuticals (medical supplies)?																			

//                  Lab reagents and test kits?																				
                if(null==workSheetGeneral.getRow(18).getCell(12).getCellTypeEnum()){
                    CM_lab = workSheetGeneral.getRow(18).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(18).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        CM_lab = ""+workSheetGeneral.getRow(18).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        CM_lab = workSheetGeneral.getRow(18).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        CM_lab = ""+workSheetGeneral.getRow(18).getCell(12).getNumericCellValue();
                        break;
                    default:
                        CM_lab = workSheetGeneral.getRow(18).getCell(12).getRawValue();
                        break;
                }
//        end  Lab reagents and test kits?																			

//                  Has the facility received supportive supervision related to commodity management in the last 3 months?																			
                if(null==workSheetGeneral.getRow(19).getCell(12).getCellTypeEnum()){
                    supervision_CM = workSheetGeneral.getRow(19).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(19).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        supervision_CM = ""+workSheetGeneral.getRow(19).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        supervision_CM = workSheetGeneral.getRow(19).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        supervision_CM = ""+workSheetGeneral.getRow(19).getCell(12).getNumericCellValue();
                        break;
                    default:
                        supervision_CM = workSheetGeneral.getRow(19).getCell(12).getRawValue();
                        break;
                }
                if(supervision_CM!=null && supervision_CM.equalsIgnoreCase("Yes")){
                    supervision_CM="1";
                }
                else{
                   supervision_CM="0";  
                }
//        end  Has the facility received supportive supervision related to commodity management in the last 3 months?																			

//                 If yes above who provided this supportive supervision?																			
                if(null==workSheetGeneral.getRow(20).getCell(12).getCellTypeEnum()){
                    who_supervision_CM = workSheetGeneral.getRow(20).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(20).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        who_supervision_CM = ""+workSheetGeneral.getRow(20).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        who_supervision_CM = workSheetGeneral.getRow(20).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        who_supervision_CM = ""+workSheetGeneral.getRow(20).getCell(12).getNumericCellValue();
                        break;
                    default:
                        who_supervision_CM = workSheetGeneral.getRow(20).getCell(12).getRawValue();
                        break;
                }
//        end  If yes above who provided this supportive supervision?																			

//                 What power source does the facility use (Electricity, Solar, etc.)?																			
                if(null==workSheetGeneral.getRow(21).getCell(12).getCellTypeEnum()){
                    power_source = workSheetGeneral.getRow(21).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(21).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        power_source = ""+workSheetGeneral.getRow(21).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        power_source = workSheetGeneral.getRow(21).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        power_source = ""+workSheetGeneral.getRow(21).getCell(12).getNumericCellValue();
                        break;
                    default:
                        power_source = workSheetGeneral.getRow(21).getCell(12).getRawValue();
                        break;
                }
//        end  What power source does the facility use (Electricity, Solar, etc.)?																			

//                 Is the power supply reliable (not more than one outage of >5 minutes per day)?																			
                if(null==workSheetGeneral.getRow(22).getCell(12).getCellTypeEnum()){
                    power_reliable = workSheetGeneral.getRow(22).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(22).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        power_reliable = ""+workSheetGeneral.getRow(22).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        power_reliable = workSheetGeneral.getRow(22).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        power_reliable = ""+workSheetGeneral.getRow(22).getCell(12).getNumericCellValue();
                        break;
                    default:
                        power_reliable = workSheetGeneral.getRow(22).getCell(12).getRawValue();
                        break;
                }
                 if(power_reliable!=null && power_reliable.equalsIgnoreCase("Yes")){
                    power_reliable="1";
                }
                else{
                   power_reliable="0";  
                }
//        end  Is the power supply reliable (not more than one outage of >5 minutes per day)?																			


//                 What are the main commodity related challenges that the facility experiences?																													
                if(null==workSheetGeneral.getRow(23).getCell(12).getCellTypeEnum()){
                    commodity_challenges = workSheetGeneral.getRow(23).getCell(12).getRawValue();
                }
                else switch (workSheetGeneral.getRow(23).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        commodity_challenges = ""+workSheetGeneral.getRow(23).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        commodity_challenges = workSheetGeneral.getRow(23).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        commodity_challenges = ""+workSheetGeneral.getRow(23).getCell(12).getNumericCellValue();
                        break;
                    default:
                        commodity_challenges = workSheetGeneral.getRow(23).getCell(12).getRawValue();
                        break;
                }
//        end  What are the main commodity related challenges that the facility experiences?																												



//   ************************************************************************


//*************SUPPLY CHAIN FACILITY PERFORMANCE SUMMARY*********************
 //              Storage Areas																													
                if(null==worksheetSummary.getRow(5).getCell(5).getCellTypeEnum()){
                    storage_pharm_num = worksheetSummary.getRow(5).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_pharm_num = ""+worksheetSummary.getRow(5).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        storage_pharm_num = worksheetSummary.getRow(5).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_pharm_num = ""+worksheetSummary.getRow(5).getCell(5).getNumericCellValue();
                        break;
                    default:
                        storage_pharm_num = worksheetSummary.getRow(5).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(6).getCellTypeEnum()){
                    storage_pharm_den = worksheetSummary.getRow(5).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_pharm_den = ""+worksheetSummary.getRow(5).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        storage_pharm_den = worksheetSummary.getRow(5).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_pharm_den = ""+worksheetSummary.getRow(5).getCell(6).getNumericCellValue();
                        break;
                    default:
                        storage_pharm_den = worksheetSummary.getRow(5).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(7).getCellTypeEnum()){
                    storage_pharm_score = worksheetSummary.getRow(5).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_pharm_score = ""+worksheetSummary.getRow(5).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        storage_pharm_score = worksheetSummary.getRow(5).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_pharm_score = ""+worksheetSummary.getRow(5).getCell(7).getNumericCellValue();
                        break;
                    default:
                        storage_pharm_score = worksheetSummary.getRow(5).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(8).getCellTypeEnum()){
                    storage_lab_num = worksheetSummary.getRow(5).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_lab_num = ""+worksheetSummary.getRow(5).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        storage_lab_num = worksheetSummary.getRow(5).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_lab_num = ""+worksheetSummary.getRow(5).getCell(8).getNumericCellValue();
                        break;
                    default:
                        storage_lab_num = worksheetSummary.getRow(5).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(9).getCellTypeEnum()){
                    storage_lab_den = worksheetSummary.getRow(5).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_lab_den = ""+worksheetSummary.getRow(5).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        storage_lab_den = worksheetSummary.getRow(5).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_lab_den = ""+worksheetSummary.getRow(5).getCell(9).getNumericCellValue();
                        break;
                    default:
                        storage_lab_den = worksheetSummary.getRow(5).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(10).getCellTypeEnum()){
                    storage_lab_score = worksheetSummary.getRow(5).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_lab_score = ""+worksheetSummary.getRow(5).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        storage_lab_score = worksheetSummary.getRow(5).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_lab_score = ""+worksheetSummary.getRow(5).getCell(10).getNumericCellValue();
                        break;
                    default:
                        storage_lab_score = worksheetSummary.getRow(5).getCell(10).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(11).getCellTypeEnum()){
                    storage_total_num = worksheetSummary.getRow(5).getCell(11).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(11).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_total_num = ""+worksheetSummary.getRow(5).getCell(11).getNumericCellValue();
                        break;
                    case STRING:
                        storage_total_num = worksheetSummary.getRow(5).getCell(11).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_total_num = ""+worksheetSummary.getRow(5).getCell(11).getNumericCellValue();
                        break;
                    default:
                        storage_total_num = worksheetSummary.getRow(5).getCell(11).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(12).getCellTypeEnum()){
                    storage_total_den = worksheetSummary.getRow(5).getCell(12).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_total_den = ""+worksheetSummary.getRow(5).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        storage_total_den = worksheetSummary.getRow(5).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_total_den = ""+worksheetSummary.getRow(5).getCell(12).getNumericCellValue();
                        break;
                    default:
                        storage_total_den = worksheetSummary.getRow(5).getCell(12).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(5).getCell(13).getCellTypeEnum()){
                    storage_total_score = worksheetSummary.getRow(5).getCell(13).getRawValue();
                }
                else switch (worksheetSummary.getRow(5).getCell(13).getCellTypeEnum()) {
                    case NUMERIC:
                        storage_total_score = ""+worksheetSummary.getRow(5).getCell(13).getNumericCellValue();
                        break;
                    case STRING:
                        storage_total_score = worksheetSummary.getRow(5).getCell(13).getStringCellValue();
                        break;
                    case FORMULA:
                        storage_total_score = ""+worksheetSummary.getRow(5).getCell(13).getNumericCellValue();
                        break;
                    default:
                        storage_total_score = worksheetSummary.getRow(5).getCell(13).getRawValue();
                        break;
                }
 //              inventory Management																													
                if(null==worksheetSummary.getRow(6).getCell(5).getCellTypeEnum()){
                    inventory_pharm_num = worksheetSummary.getRow(6).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_pharm_num = ""+worksheetSummary.getRow(6).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_pharm_num = worksheetSummary.getRow(6).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_pharm_num = ""+worksheetSummary.getRow(6).getCell(5).getNumericCellValue();
                        break;
                    default:
                        inventory_pharm_num = worksheetSummary.getRow(6).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(6).getCellTypeEnum()){
                    inventory_pharm_den = worksheetSummary.getRow(6).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_pharm_den = ""+worksheetSummary.getRow(6).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_pharm_den = worksheetSummary.getRow(6).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_pharm_den = ""+worksheetSummary.getRow(6).getCell(6).getNumericCellValue();
                        break;
                    default:
                        inventory_pharm_den = worksheetSummary.getRow(6).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(7).getCellTypeEnum()){
                    inventory_pharm_score = worksheetSummary.getRow(6).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_pharm_score = ""+worksheetSummary.getRow(6).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_pharm_score = worksheetSummary.getRow(6).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_pharm_score = ""+worksheetSummary.getRow(6).getCell(7).getNumericCellValue();
                        break;
                    default:
                        inventory_pharm_score = worksheetSummary.getRow(6).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(8).getCellTypeEnum()){
                    inventory_lab_num = worksheetSummary.getRow(6).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_lab_num = ""+worksheetSummary.getRow(6).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_lab_num = worksheetSummary.getRow(6).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_lab_num = ""+worksheetSummary.getRow(6).getCell(8).getNumericCellValue();
                        break;
                    default:
                        inventory_lab_num = worksheetSummary.getRow(6).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(9).getCellTypeEnum()){
                    inventory_lab_den = worksheetSummary.getRow(6).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_lab_den = ""+worksheetSummary.getRow(6).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_lab_den = worksheetSummary.getRow(6).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_lab_den = ""+worksheetSummary.getRow(6).getCell(9).getNumericCellValue();
                        break;
                    default:
                        inventory_lab_den = worksheetSummary.getRow(6).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(10).getCellTypeEnum()){
                    inventory_lab_score = worksheetSummary.getRow(6).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_lab_score = ""+worksheetSummary.getRow(6).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_lab_score = worksheetSummary.getRow(6).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_lab_score = ""+worksheetSummary.getRow(6).getCell(10).getNumericCellValue();
                        break;
                    default:
                        inventory_lab_score = worksheetSummary.getRow(6).getCell(10).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(11).getCellTypeEnum()){
                    inventory_total_num = worksheetSummary.getRow(6).getCell(11).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(11).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_total_num = ""+worksheetSummary.getRow(6).getCell(11).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_total_num = worksheetSummary.getRow(6).getCell(11).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_total_num = ""+worksheetSummary.getRow(6).getCell(11).getNumericCellValue();
                        break;
                    default:
                        inventory_total_num = worksheetSummary.getRow(6).getCell(11).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(12).getCellTypeEnum()){
                    inventory_total_den = worksheetSummary.getRow(6).getCell(12).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_total_den = ""+worksheetSummary.getRow(6).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_total_den = worksheetSummary.getRow(6).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_total_den = ""+worksheetSummary.getRow(6).getCell(12).getNumericCellValue();
                        break;
                    default:
                        inventory_total_den = worksheetSummary.getRow(6).getCell(12).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(6).getCell(13).getCellTypeEnum()){
                    inventory_total_score = worksheetSummary.getRow(6).getCell(13).getRawValue();
                }
                else switch (worksheetSummary.getRow(6).getCell(13).getCellTypeEnum()) {
                    case NUMERIC:
                        inventory_total_score = ""+worksheetSummary.getRow(6).getCell(13).getNumericCellValue();
                        break;
                    case STRING:
                        inventory_total_score = worksheetSummary.getRow(6).getCell(13).getStringCellValue();
                        break;
                    case FORMULA:
                        inventory_total_score = ""+worksheetSummary.getRow(6).getCell(13).getNumericCellValue();
                        break;
                    default:
                        inventory_total_score = worksheetSummary.getRow(6).getCell(13).getRawValue();
                        break;
                }
 //              Resources & Reference Materials																													
                if(null==worksheetSummary.getRow(7).getCell(5).getCellTypeEnum()){
                    RRM_pharm_num = worksheetSummary.getRow(7).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_pharm_num = ""+worksheetSummary.getRow(7).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_pharm_num = worksheetSummary.getRow(7).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_pharm_num = ""+worksheetSummary.getRow(7).getCell(5).getNumericCellValue();
                        break;
                    default:
                        RRM_pharm_num = worksheetSummary.getRow(7).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(6).getCellTypeEnum()){
                    RRM_pharm_den = worksheetSummary.getRow(7).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_pharm_den = ""+worksheetSummary.getRow(7).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_pharm_den = worksheetSummary.getRow(7).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_pharm_den = ""+worksheetSummary.getRow(7).getCell(6).getNumericCellValue();
                        break;
                    default:
                        RRM_pharm_den = worksheetSummary.getRow(7).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(7).getCellTypeEnum()){
                    RRM_pharm_score = worksheetSummary.getRow(7).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_pharm_score = ""+worksheetSummary.getRow(7).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_pharm_score = worksheetSummary.getRow(7).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_pharm_score = ""+worksheetSummary.getRow(7).getCell(7).getNumericCellValue();
                        break;
                    default:
                        RRM_pharm_score = worksheetSummary.getRow(7).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(8).getCellTypeEnum()){
                    RRM_lab_num = worksheetSummary.getRow(7).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_lab_num = ""+worksheetSummary.getRow(7).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_lab_num = worksheetSummary.getRow(7).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_lab_num = ""+worksheetSummary.getRow(7).getCell(8).getNumericCellValue();
                        break;
                    default:
                        RRM_lab_num = worksheetSummary.getRow(7).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(9).getCellTypeEnum()){
                    RRM_lab_den = worksheetSummary.getRow(7).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_lab_den = ""+worksheetSummary.getRow(7).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_lab_den = worksheetSummary.getRow(7).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_lab_den = ""+worksheetSummary.getRow(7).getCell(9).getNumericCellValue();
                        break;
                    default:
                        RRM_lab_den = worksheetSummary.getRow(7).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(10).getCellTypeEnum()){
                    RRM_lab_score = worksheetSummary.getRow(7).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_lab_score = ""+worksheetSummary.getRow(7).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_lab_score = worksheetSummary.getRow(7).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_lab_score = ""+worksheetSummary.getRow(7).getCell(10).getNumericCellValue();
                        break;
                    default:
                        RRM_lab_score = worksheetSummary.getRow(7).getCell(10).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(11).getCellTypeEnum()){
                    RRM_total_num = worksheetSummary.getRow(7).getCell(11).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(11).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_total_num = ""+worksheetSummary.getRow(7).getCell(11).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_total_num = worksheetSummary.getRow(7).getCell(11).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_total_num = ""+worksheetSummary.getRow(7).getCell(11).getNumericCellValue();
                        break;
                    default:
                        RRM_total_num = worksheetSummary.getRow(7).getCell(11).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(12).getCellTypeEnum()){
                    RRM_total_den = worksheetSummary.getRow(7).getCell(12).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_total_den = ""+worksheetSummary.getRow(7).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_total_den = worksheetSummary.getRow(7).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_total_den = ""+worksheetSummary.getRow(7).getCell(12).getNumericCellValue();
                        break;
                    default:
                        RRM_total_den = worksheetSummary.getRow(7).getCell(12).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(7).getCell(13).getCellTypeEnum()){
                    RRM_total_score = worksheetSummary.getRow(7).getCell(13).getRawValue();
                }
                else switch (worksheetSummary.getRow(7).getCell(13).getCellTypeEnum()) {
                    case NUMERIC:
                        RRM_total_score = ""+worksheetSummary.getRow(7).getCell(13).getNumericCellValue();
                        break;
                    case STRING:
                        RRM_total_score = worksheetSummary.getRow(7).getCell(13).getStringCellValue();
                        break;
                    case FORMULA:
                        RRM_total_score = ""+worksheetSummary.getRow(7).getCell(13).getNumericCellValue();
                        break;
                    default:
                        RRM_total_score = worksheetSummary.getRow(7).getCell(13).getRawValue();
                        break;
                }
                
                
                
 //              Availability and Use of MIS Tools																													
                if(null==worksheetSummary.getRow(8).getCell(5).getCellTypeEnum()){
                    MIS_pharm_num = worksheetSummary.getRow(8).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_pharm_num = ""+worksheetSummary.getRow(8).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_pharm_num = worksheetSummary.getRow(8).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_pharm_num = ""+worksheetSummary.getRow(8).getCell(5).getNumericCellValue();
                        break;
                    default:
                        MIS_pharm_num = worksheetSummary.getRow(8).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(6).getCellTypeEnum()){
                    MIS_pharm_den = worksheetSummary.getRow(8).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_pharm_den = ""+worksheetSummary.getRow(8).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_pharm_den = worksheetSummary.getRow(8).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_pharm_den = ""+worksheetSummary.getRow(8).getCell(6).getNumericCellValue();
                        break;
                    default:
                        MIS_pharm_den = worksheetSummary.getRow(8).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(7).getCellTypeEnum()){
                    MIS_pharm_score = worksheetSummary.getRow(8).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_pharm_score = ""+worksheetSummary.getRow(8).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_pharm_score = worksheetSummary.getRow(8).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_pharm_score = ""+worksheetSummary.getRow(8).getCell(7).getNumericCellValue();
                        break;
                    default:
                        MIS_pharm_score = worksheetSummary.getRow(8).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(8).getCellTypeEnum()){
                    MIS_lab_num = worksheetSummary.getRow(8).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_lab_num = ""+worksheetSummary.getRow(8).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_lab_num = worksheetSummary.getRow(8).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_lab_num = ""+worksheetSummary.getRow(8).getCell(8).getNumericCellValue();
                        break;
                    default:
                        MIS_lab_num = worksheetSummary.getRow(8).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(9).getCellTypeEnum()){
                    MIS_lab_den = worksheetSummary.getRow(8).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_lab_den = ""+worksheetSummary.getRow(8).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_lab_den = worksheetSummary.getRow(8).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_lab_den = ""+worksheetSummary.getRow(8).getCell(9).getNumericCellValue();
                        break;
                    default:
                        MIS_lab_den = worksheetSummary.getRow(8).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(10).getCellTypeEnum()){
                    MIS_lab_score = worksheetSummary.getRow(8).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_lab_score = ""+worksheetSummary.getRow(8).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_lab_score = worksheetSummary.getRow(8).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_lab_score = ""+worksheetSummary.getRow(8).getCell(10).getNumericCellValue();
                        break;
                    default:
                        MIS_lab_score = worksheetSummary.getRow(8).getCell(10).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(11).getCellTypeEnum()){
                    MIS_total_num = worksheetSummary.getRow(8).getCell(11).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(11).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_total_num = ""+worksheetSummary.getRow(8).getCell(11).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_total_num = worksheetSummary.getRow(8).getCell(11).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_total_num = ""+worksheetSummary.getRow(8).getCell(11).getNumericCellValue();
                        break;
                    default:
                        MIS_total_num = worksheetSummary.getRow(8).getCell(11).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(12).getCellTypeEnum()){
                    MIS_total_den = worksheetSummary.getRow(8).getCell(12).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_total_den = ""+worksheetSummary.getRow(8).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_total_den = worksheetSummary.getRow(8).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_total_den = ""+worksheetSummary.getRow(8).getCell(12).getNumericCellValue();
                        break;
                    default:
                        MIS_total_den = worksheetSummary.getRow(8).getCell(12).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(8).getCell(13).getCellTypeEnum()){
                    MIS_total_score = worksheetSummary.getRow(8).getCell(13).getRawValue();
                }
                else switch (worksheetSummary.getRow(8).getCell(13).getCellTypeEnum()) {
                    case NUMERIC:
                        MIS_total_score = ""+worksheetSummary.getRow(8).getCell(13).getNumericCellValue();
                        break;
                    case STRING:
                        MIS_total_score = worksheetSummary.getRow(8).getCell(13).getStringCellValue();
                        break;
                    case FORMULA:
                        MIS_total_score = ""+worksheetSummary.getRow(8).getCell(13).getNumericCellValue();
                        break;
                    default:
                        MIS_total_score = worksheetSummary.getRow(8).getCell(13).getRawValue();
                        break;
                }
                
                
 //              Aggregate Score																												
                if(null==worksheetSummary.getRow(9).getCell(5).getCellTypeEnum()){
                    total_pharm_num = worksheetSummary.getRow(9).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        total_pharm_num = ""+worksheetSummary.getRow(9).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        total_pharm_num = worksheetSummary.getRow(9).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        total_pharm_num = ""+worksheetSummary.getRow(9).getCell(5).getNumericCellValue();
                        break;
                    default:
                        total_pharm_num = worksheetSummary.getRow(9).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(6).getCellTypeEnum()){
                    total_pharm_den = worksheetSummary.getRow(9).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        total_pharm_den = ""+worksheetSummary.getRow(9).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        total_pharm_den = worksheetSummary.getRow(9).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        total_pharm_den = ""+worksheetSummary.getRow(9).getCell(6).getNumericCellValue();
                        break;
                    default:
                        total_pharm_den = worksheetSummary.getRow(9).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(7).getCellTypeEnum()){
                    total_pharm_score = worksheetSummary.getRow(9).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        total_pharm_score = ""+worksheetSummary.getRow(9).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        total_pharm_score = worksheetSummary.getRow(9).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        total_pharm_score = ""+worksheetSummary.getRow(9).getCell(7).getNumericCellValue();
                        break;
                    default:
                        total_pharm_score = worksheetSummary.getRow(9).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(8).getCellTypeEnum()){
                    total_lab_num = worksheetSummary.getRow(9).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        total_lab_num = ""+worksheetSummary.getRow(9).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        total_lab_num = worksheetSummary.getRow(9).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        total_lab_num = ""+worksheetSummary.getRow(9).getCell(8).getNumericCellValue();
                        break;
                    default:
                        total_lab_num = worksheetSummary.getRow(9).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(9).getCellTypeEnum()){
                    total_lab_den = worksheetSummary.getRow(9).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        total_lab_den = ""+worksheetSummary.getRow(9).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        total_lab_den = worksheetSummary.getRow(9).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        total_lab_den = ""+worksheetSummary.getRow(9).getCell(9).getNumericCellValue();
                        break;
                    default:
                        total_lab_den = worksheetSummary.getRow(9).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(10).getCellTypeEnum()){
                    total_lab_score = worksheetSummary.getRow(9).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        total_lab_score = ""+worksheetSummary.getRow(9).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        total_lab_score = worksheetSummary.getRow(9).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        total_lab_score = ""+worksheetSummary.getRow(9).getCell(10).getNumericCellValue();
                        break;
                    default:
                        total_lab_score = worksheetSummary.getRow(9).getCell(10).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(11).getCellTypeEnum()){
                    total_num = worksheetSummary.getRow(9).getCell(11).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(11).getCellTypeEnum()) {
                    case NUMERIC:
                        total_num = ""+worksheetSummary.getRow(9).getCell(11).getNumericCellValue();
                        break;
                    case STRING:
                        total_num = worksheetSummary.getRow(9).getCell(11).getStringCellValue();
                        break;
                    case FORMULA:
                        total_num = ""+worksheetSummary.getRow(9).getCell(11).getNumericCellValue();
                        break;
                    default:
                        total_num = worksheetSummary.getRow(9).getCell(11).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(12).getCellTypeEnum()){
                    total_den = worksheetSummary.getRow(9).getCell(12).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(12).getCellTypeEnum()) {
                    case NUMERIC:
                        total_den = ""+worksheetSummary.getRow(9).getCell(12).getNumericCellValue();
                        break;
                    case STRING:
                        total_den = worksheetSummary.getRow(9).getCell(12).getStringCellValue();
                        break;
                    case FORMULA:
                        total_den = ""+worksheetSummary.getRow(9).getCell(12).getNumericCellValue();
                        break;
                    default:
                        total_den = worksheetSummary.getRow(9).getCell(12).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(9).getCell(13).getCellTypeEnum()){
                    total_score = worksheetSummary.getRow(9).getCell(13).getRawValue();
                }
                else switch (worksheetSummary.getRow(9).getCell(13).getCellTypeEnum()) {
                    case NUMERIC:
                        total_score = ""+worksheetSummary.getRow(9).getCell(13).getNumericCellValue();
                        break;
                    case STRING:
                        total_score = worksheetSummary.getRow(9).getCell(13).getStringCellValue();
                        break;
                    case FORMULA:
                        total_score = ""+worksheetSummary.getRow(9).getCell(13).getNumericCellValue();
                        break;
                    default:
                        total_score = worksheetSummary.getRow(9).getCell(13).getRawValue();
                        break;
                }
                
              
                                
                
 //              HIV RTKs End User Verification																													
                if(null==worksheetSummary.getRow(11).getCell(5).getCellTypeEnum()){
                    EUV_pharm_num = worksheetSummary.getRow(11).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(11).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        EUV_pharm_num = ""+worksheetSummary.getRow(11).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        EUV_pharm_num = worksheetSummary.getRow(11).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        EUV_pharm_num = ""+worksheetSummary.getRow(11).getCell(5).getNumericCellValue();
                        break;
                    default:
                        EUV_pharm_num = worksheetSummary.getRow(11).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(11).getCell(6).getCellTypeEnum()){
                    EUV_pharm_den = worksheetSummary.getRow(11).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(11).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        EUV_pharm_den = ""+worksheetSummary.getRow(11).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        EUV_pharm_den = worksheetSummary.getRow(11).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        EUV_pharm_den = ""+worksheetSummary.getRow(11).getCell(6).getNumericCellValue();
                        break;
                    default:
                        EUV_pharm_den = worksheetSummary.getRow(11).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(11).getCell(7).getCellTypeEnum()){
                    EUV_pharm_score = worksheetSummary.getRow(11).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(11).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        EUV_pharm_score = ""+worksheetSummary.getRow(11).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        EUV_pharm_score = worksheetSummary.getRow(11).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        EUV_pharm_score = ""+worksheetSummary.getRow(11).getCell(7).getNumericCellValue();
                        break;
                    default:
                        EUV_pharm_score = worksheetSummary.getRow(11).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(11).getCell(8).getCellTypeEnum()){
                    EUV_lab_num = worksheetSummary.getRow(11).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(11).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        EUV_lab_num = ""+worksheetSummary.getRow(11).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        EUV_lab_num = worksheetSummary.getRow(11).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        EUV_lab_num = ""+worksheetSummary.getRow(11).getCell(8).getNumericCellValue();
                        break;
                    default:
                        EUV_lab_num = worksheetSummary.getRow(11).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(11).getCell(9).getCellTypeEnum()){
                    EUV_lab_den = worksheetSummary.getRow(11).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(11).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        EUV_lab_den = ""+worksheetSummary.getRow(11).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        EUV_lab_den = worksheetSummary.getRow(11).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        EUV_lab_den = ""+worksheetSummary.getRow(11).getCell(9).getNumericCellValue();
                        break;
                    default:
                        EUV_lab_den = worksheetSummary.getRow(11).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(11).getCell(10).getCellTypeEnum()){
                    EUV_lab_score = worksheetSummary.getRow(11).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(11).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        EUV_lab_score = ""+worksheetSummary.getRow(11).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        EUV_lab_score = worksheetSummary.getRow(11).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        EUV_lab_score = ""+worksheetSummary.getRow(11).getCell(10).getNumericCellValue();
                        break;
                    default:
                        EUV_lab_score = worksheetSummary.getRow(11).getCell(10).getRawValue();
                        break;
                }

                
 //              Inventory Management: additional commodities																												
                if(null==worksheetSummary.getRow(12).getCell(5).getCellTypeEnum()){
                    IM_additional_pharm_num = worksheetSummary.getRow(12).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(12).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        IM_additional_pharm_num = ""+worksheetSummary.getRow(12).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        IM_additional_pharm_num = worksheetSummary.getRow(12).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        IM_additional_pharm_num = ""+worksheetSummary.getRow(12).getCell(5).getNumericCellValue();
                        break;
                    default:
                        IM_additional_pharm_num = worksheetSummary.getRow(12).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(12).getCell(6).getCellTypeEnum()){
                    IM_additional_pharm_den = worksheetSummary.getRow(12).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(12).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        IM_additional_pharm_den = ""+worksheetSummary.getRow(12).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        IM_additional_pharm_den = worksheetSummary.getRow(12).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        IM_additional_pharm_den = ""+worksheetSummary.getRow(12).getCell(6).getNumericCellValue();
                        break;
                    default:
                        IM_additional_pharm_den = worksheetSummary.getRow(12).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(12).getCell(7).getCellTypeEnum()){
                    IM_additional_pharm_score = worksheetSummary.getRow(12).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(12).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        IM_additional_pharm_score = ""+worksheetSummary.getRow(12).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        IM_additional_pharm_score = worksheetSummary.getRow(12).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        IM_additional_pharm_score = ""+worksheetSummary.getRow(12).getCell(7).getNumericCellValue();
                        break;
                    default:
                        IM_additional_pharm_score = worksheetSummary.getRow(12).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(12).getCell(8).getCellTypeEnum()){
                    IM_additional_lab_num = worksheetSummary.getRow(12).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(12).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        IM_additional_lab_num = ""+worksheetSummary.getRow(12).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        IM_additional_lab_num = worksheetSummary.getRow(12).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        IM_additional_lab_num = ""+worksheetSummary.getRow(12).getCell(8).getNumericCellValue();
                        break;
                    default:
                        IM_additional_lab_num = worksheetSummary.getRow(12).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(12).getCell(9).getCellTypeEnum()){
                    IM_additional_lab_den = worksheetSummary.getRow(12).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(12).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        IM_additional_lab_den = ""+worksheetSummary.getRow(12).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        IM_additional_lab_den = worksheetSummary.getRow(12).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        IM_additional_lab_den = ""+worksheetSummary.getRow(12).getCell(9).getNumericCellValue();
                        break;
                    default:
                        IM_additional_lab_den = worksheetSummary.getRow(12).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(12).getCell(10).getCellTypeEnum()){
                    IM_additional_lab_score = worksheetSummary.getRow(12).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(12).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        IM_additional_lab_score = ""+worksheetSummary.getRow(12).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        IM_additional_lab_score = worksheetSummary.getRow(12).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        IM_additional_lab_score = ""+worksheetSummary.getRow(12).getCell(10).getNumericCellValue();
                        break;
                    default:
                        IM_additional_lab_score = worksheetSummary.getRow(12).getCell(10).getRawValue();
                        break;
                }

                

                
 //              Inventory Management: additional commodities																												
                if(null==worksheetSummary.getRow(13).getCell(5).getCellTypeEnum()){
                    MTC_pharm_num = worksheetSummary.getRow(13).getCell(5).getRawValue();
                }
                else switch (worksheetSummary.getRow(13).getCell(5).getCellTypeEnum()) {
                    case NUMERIC:
                        MTC_pharm_num = ""+worksheetSummary.getRow(13).getCell(5).getNumericCellValue();
                        break;
                    case STRING:
                        MTC_pharm_num = worksheetSummary.getRow(13).getCell(5).getStringCellValue();
                        break;
                    case FORMULA:
                        MTC_pharm_num = ""+worksheetSummary.getRow(13).getCell(5).getNumericCellValue();
                        break;
                    default:
                        MTC_pharm_num = worksheetSummary.getRow(13).getCell(5).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(13).getCell(6).getCellTypeEnum()){
                    MTC_pharm_den = worksheetSummary.getRow(13).getCell(6).getRawValue();
                }
                else switch (worksheetSummary.getRow(13).getCell(6).getCellTypeEnum()) {
                    case NUMERIC:
                        MTC_pharm_den = ""+worksheetSummary.getRow(13).getCell(6).getNumericCellValue();
                        break;
                    case STRING:
                        MTC_pharm_den = worksheetSummary.getRow(13).getCell(6).getStringCellValue();
                        break;
                    case FORMULA:
                        MTC_pharm_den = ""+worksheetSummary.getRow(13).getCell(6).getNumericCellValue();
                        break;
                    default:
                        MTC_pharm_den = worksheetSummary.getRow(13).getCell(6).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(13).getCell(7).getCellTypeEnum()){
                    MTC_pharm_score = worksheetSummary.getRow(13).getCell(7).getRawValue();
                }
                else switch (worksheetSummary.getRow(13).getCell(7).getCellTypeEnum()) {
                    case NUMERIC:
                        MTC_pharm_score = ""+worksheetSummary.getRow(13).getCell(7).getNumericCellValue();
                        break;
                    case STRING:
                        MTC_pharm_score = worksheetSummary.getRow(13).getCell(7).getStringCellValue();
                        break;
                    case FORMULA:
                        MTC_pharm_score = ""+worksheetSummary.getRow(13).getCell(7).getNumericCellValue();
                        break;
                    default:
                        MTC_pharm_score = worksheetSummary.getRow(13).getCell(7).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(13).getCell(8).getCellTypeEnum()){
                    MTC_lab_num = worksheetSummary.getRow(13).getCell(8).getRawValue();
                }
                else switch (worksheetSummary.getRow(13).getCell(8).getCellTypeEnum()) {
                    case NUMERIC:
                        MTC_lab_num = ""+worksheetSummary.getRow(13).getCell(8).getNumericCellValue();
                        break;
                    case STRING:
                        MTC_lab_num = worksheetSummary.getRow(13).getCell(8).getStringCellValue();
                        break;
                    case FORMULA:
                        MTC_lab_num = ""+worksheetSummary.getRow(13).getCell(8).getNumericCellValue();
                        break;
                    default:
                        MTC_lab_num = worksheetSummary.getRow(13).getCell(8).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(13).getCell(9).getCellTypeEnum()){
                    MTC_lab_den = worksheetSummary.getRow(13).getCell(9).getRawValue();
                }
                else switch (worksheetSummary.getRow(13).getCell(9).getCellTypeEnum()) {
                    case NUMERIC:
                        MTC_lab_den = ""+worksheetSummary.getRow(13).getCell(9).getNumericCellValue();
                        break;
                    case STRING:
                        MTC_lab_den = worksheetSummary.getRow(13).getCell(9).getStringCellValue();
                        break;
                    case FORMULA:
                        MTC_lab_den = ""+worksheetSummary.getRow(13).getCell(9).getNumericCellValue();
                        break;
                    default:
                        MTC_lab_den = worksheetSummary.getRow(13).getCell(9).getRawValue();
                        break;
                }
                if(null==worksheetSummary.getRow(13).getCell(10).getCellTypeEnum()){
                    MTC_lab_score = worksheetSummary.getRow(13).getCell(10).getRawValue();
                }
                else switch (worksheetSummary.getRow(13).getCell(10).getCellTypeEnum()) {
                    case NUMERIC:
                        MTC_lab_score = ""+worksheetSummary.getRow(13).getCell(10).getNumericCellValue();
                        break;
                    case STRING:
                        MTC_lab_score = worksheetSummary.getRow(13).getCell(10).getStringCellValue();
                        break;
                    case FORMULA:
                        MTC_lab_score = ""+worksheetSummary.getRow(13).getCell(10).getNumericCellValue();
                        break;
                    default:
                        MTC_lab_score = worksheetSummary.getRow(13).getCell(10).getRawValue();
                        break;
                }

                
//       END OF READING VALUES

if(mfl_code!=null){
if(!mfl_code.equals("")){
//checker 
mfl_code = mfl_code.replace(".0", "");
//GET CORRECT COUNTY, SUBCOUNTY DATA FOR THIS FACILITY
String getFacilDetails="SELECT County, DistrictNom,SubPartnerNom " +
"FROM  subpartnera join district on subpartnera.DistrictID=district.DistrictID " +
"join county on county.CountyID=district.CountyID " +
"where subpartnera.CentreSanteId=?";
conn.pst2 = conn.conn.prepareStatement(getFacilDetails);
conn.pst2.setString(1, mfl_code);
conn.rs2 = conn.pst2.executeQuery();
if(conn.rs2.next()){
  county = conn.rs2.getString(1);
  sub_county = conn.rs2.getString(2);
  facility = conn.rs2.getString(3);
          
String checker = "SELECT id FROM report WHERE mfl_code=?";
conn.pst = conn.conn.prepareStatement(checker);
conn.pst.setString(1, mfl_code);
conn.rs = conn.pst.executeQuery();
                System.out.println("checker query : "+conn.pst);
if(conn.rs.next()){
    //update record
    id=conn.rs.getString(1);
    
    query = "UPDATE report SET facility=?,facility_level=?,county=?,sub_county=?,mfl_code=?,level=?,ownership=?,facility_incharge=?,incharge_contact=?,pharmacy_person=?,pharmacy_phone=?,lab_person=?,lab_phone=?,visit_date=?,has_lab=?,has_designated_store=?,CM_pharmaceutical=?,CM_non_pharmaceutical=?,CM_lab=?,supervision_CM=?,who_supervision_CM=?,power_source=?,power_reliable=?,commodity_challenges=?,storage_pharm_num=?,storage_pharm_den=?,storage_pharm_score=?,storage_lab_num=?,storage_lab_den=?,storage_lab_score=?,storage_total_num=?,storage_total_den=?,storage_total_score=?,inventory_pharm_num=?,inventory_pharm_den=?,inventory_pharm_score=?,inventory_lab_num=?,inventory_lab_den=?,inventory_lab_score=?,inventory_total_num=?,inventory_total_den=?,inventory_total_score=?,RRM_pharm_num=?,RRM_pharm_den=?,RRM_pharm_score=?,RRM_lab_num=?,RRM_lab_den=?,RRM_lab_score=?,RRM_total_num=?,RRM_total_den=?,RRM_total_score=?,MIS_pharm_num=?,MIS_pharm_den=?,MIS_pharm_score=?,MIS_lab_num=?,MIS_lab_den=?,MIS_lab_score=?,MIS_total_num=?,MIS_total_den=?,MIS_total_score=?,total_pharm_num=?,total_pharm_den=?,total_pharm_score=?,total_lab_num=?,total_lab_den=?,total_lab_score=?,total_num=?,total_den=?,total_score=?,EUV_pharm_num=?,EUV_pharm_den=?,EUV_pharm_score=?,EUV_lab_num=?,EUV_lab_den=?,EUV_lab_score=?,IM_additional_pharm_num=?,IM_additional_pharm_den=?,IM_additional_pharm_score=?,IM_additional_lab_num=?,IM_additional_lab_den=?,IM_additional_lab_score=?,MTC_pharm_num=?,MTC_pharm_den=?,MTC_pharm_score=?,MTC_lab_num=?,MTC_lab_den=?,MTC_lab_score=? WHERE id=? ";
    conn.pst1 = conn.conn.prepareStatement(query);
    conn.pst1.setString(88, id);

}
else{
    query = "INSERT INTO report (facility,facility_level,county,sub_county,mfl_code,level,ownership,facility_incharge,incharge_contact,pharmacy_person,pharmacy_phone,lab_person,lab_phone,visit_date,has_lab,has_designated_store,CM_pharmaceutical,CM_non_pharmaceutical,CM_lab,supervision_CM,who_supervision_CM,power_source,power_reliable,commodity_challenges,storage_pharm_num,storage_pharm_den,storage_pharm_score,storage_lab_num,storage_lab_den,storage_lab_score,storage_total_num,storage_total_den,storage_total_score,inventory_pharm_num,inventory_pharm_den,inventory_pharm_score,inventory_lab_num,inventory_lab_den,inventory_lab_score,inventory_total_num,inventory_total_den,inventory_total_score,RRM_pharm_num,RRM_pharm_den,RRM_pharm_score,RRM_lab_num,RRM_lab_den,RRM_lab_score,RRM_total_num,RRM_total_den,RRM_total_score,MIS_pharm_num,MIS_pharm_den,MIS_pharm_score,MIS_lab_num,MIS_lab_den,MIS_lab_score,MIS_total_num,MIS_total_den,MIS_total_score,total_pharm_num,total_pharm_den,total_pharm_score,total_lab_num,total_lab_den,total_lab_score,total_num,total_den,total_score,EUV_pharm_num,EUV_pharm_den,EUV_pharm_score,EUV_lab_num,EUV_lab_den,EUV_lab_score,IM_additional_pharm_num,IM_additional_pharm_den,IM_additional_pharm_score,IM_additional_lab_num,IM_additional_lab_den,IM_additional_lab_score,MTC_pharm_num,MTC_pharm_den,MTC_pharm_score,MTC_lab_num,MTC_lab_den,MTC_lab_score) VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?,?)";   
    conn.pst1 = conn.conn.prepareStatement(query);
}
conn.pst1.setString(1, facility);
conn.pst1.setString(2, facility_level);
conn.pst1.setString(3, county);
conn.pst1.setString(4, sub_county);
conn.pst1.setString(5, mfl_code);
conn.pst1.setString(6, level);
conn.pst1.setString(7, ownership);
conn.pst1.setString(8, facility_incharge);
conn.pst1.setString(9, incharge_contact);
conn.pst1.setString(10, pharmacy_person);
conn.pst1.setString(11, pharmacy_phone);
conn.pst1.setString(12, lab_person);
conn.pst1.setString(13, lab_phone);
conn.pst1.setString(14, visit_date);
conn.pst1.setString(15, has_lab);
conn.pst1.setString(16, has_designated_store);
conn.pst1.setString(17, CM_pharmaceutical);
conn.pst1.setString(18, CM_non_pharmaceutical);
conn.pst1.setString(19, CM_lab);
conn.pst1.setString(20, supervision_CM); 
conn.pst1.setString(21, who_supervision_CM);
conn.pst1.setString(22, power_source);
conn.pst1.setString(23, power_reliable);
conn.pst1.setString(24, commodity_challenges);
conn.pst1.setString(25, storage_pharm_num);
conn.pst1.setString(26, storage_pharm_den);
conn.pst1.setString(27, storage_pharm_score);
conn.pst1.setString(28, storage_lab_num);
conn.pst1.setString(29, storage_lab_den);
conn.pst1.setString(30, storage_lab_score);
conn.pst1.setString(31, storage_total_num);
conn.pst1.setString(32, storage_total_den);
conn.pst1.setString(33, storage_total_score);
conn.pst1.setString(34, inventory_pharm_num);
conn.pst1.setString(35, inventory_pharm_den);
conn.pst1.setString(36, inventory_pharm_score);
conn.pst1.setString(37, inventory_lab_num);
conn.pst1.setString(38, inventory_lab_den);
conn.pst1.setString(39, inventory_lab_score);
conn.pst1.setString(40, inventory_total_num);
conn.pst1.setString(41, inventory_total_den);
conn.pst1.setString(42, inventory_total_score);
conn.pst1.setString(43, RRM_pharm_num);
conn.pst1.setString(44, RRM_pharm_den);
conn.pst1.setString(45, RRM_pharm_score);
conn.pst1.setString(46, RRM_lab_num);
conn.pst1.setString(47, RRM_lab_den);
conn.pst1.setString(48, RRM_lab_score);
conn.pst1.setString(49, RRM_total_num);
conn.pst1.setString(50, RRM_total_den);
conn.pst1.setString(51, RRM_total_score);
conn.pst1.setString(52, MIS_pharm_num);
conn.pst1.setString(53, MIS_pharm_den);
conn.pst1.setString(54, MIS_pharm_score);
conn.pst1.setString(55, MIS_lab_num);
conn.pst1.setString(56, MIS_lab_den);
conn.pst1.setString(57, MIS_lab_score);
conn.pst1.setString(58, MIS_total_num);
conn.pst1.setString(59, MIS_total_den);
conn.pst1.setString(60, MIS_total_score);
conn.pst1.setString(61, total_pharm_num);
conn.pst1.setString(62, total_pharm_den);
conn.pst1.setString(63, total_pharm_score);
conn.pst1.setString(64, total_lab_num);
conn.pst1.setString(65, total_lab_den);
conn.pst1.setString(66, total_lab_score);
conn.pst1.setString(67, total_num);
conn.pst1.setString(68, total_den);
conn.pst1.setString(69, total_score);
conn.pst1.setString(70, EUV_pharm_num);
conn.pst1.setString(71, EUV_pharm_den);
conn.pst1.setString(72, EUV_pharm_score);
conn.pst1.setString(73, EUV_lab_num);
conn.pst1.setString(74, EUV_lab_den);
conn.pst1.setString(75, EUV_lab_score);
conn.pst1.setString(76, IM_additional_pharm_num);
conn.pst1.setString(77, IM_additional_pharm_den);
conn.pst1.setString(78, IM_additional_pharm_score);
conn.pst1.setString(79, IM_additional_lab_num);
conn.pst1.setString(80, IM_additional_lab_den);
conn.pst1.setString(81, IM_additional_lab_score);
conn.pst1.setString(82, MTC_pharm_num);
conn.pst1.setString(83, MTC_pharm_den);
conn.pst1.setString(84, MTC_pharm_score);
conn.pst1.setString(85, MTC_lab_num);
conn.pst1.setString(86, MTC_lab_den);
conn.pst1.setString(87, MTC_lab_score);

int num = conn.pst1.executeUpdate();
  if(num>0){
      added++;
  }  
  else{
      number_changes++;
  }
}
else{
    failed_description="No such MFL Code in our Internal System\\'s Master Facility List";
    failed++;
}
}
else{
   failed_description="Mising MFL Code. It is Blank";
   failed++;
}
}
else{
    failed_description="Mising MFL Code . It is NULL";
    failed++;
}
//***************************************************************************
    }
            if(failed>0){
                JSONObject obj = new JSONObject();
                obj.put("file_name", fileName);
                obj.put("description", failed_description);
                jarray.add(obj);
            }
        }
        
        finalobj.put("data",jarray);
        finalobj.put("uploaded",added);
        finalobj.put("no_changes",number_changes);
        System.out.println("errors : "+finalobj);
        session.setAttribute("upload_errors", finalobj);
        
        response.sendRedirect("upload_data.jsp");
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
          Logger.getLogger(upload_excels.class.getName()).log(Level.SEVERE, null, ex);
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
          Logger.getLogger(upload_excels.class.getName()).log(Level.SEVERE, null, ex);
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
 private String getFileName(Part part) {
            String file_name="";
        String contentDisp = part.getHeader("content-disposition");
//        System.out.println("content-disposition header= "+contentDisp);
        String[] tokens = contentDisp.split(";");
      
        for (String token : tokens) {
            if (token.trim().startsWith("filename")) {
                file_name = token.substring(token.indexOf("=") + 2, token.length()-1);
              break;  
            }
            
        }
//         System.out.println("content-disposition final : "+file_name);
        return file_name;
    }
 
public String removeLast(String str, int num) {
    if (str != null && str.length() > 0) {
        str = str.substring(0, str.length() - num);
    }
    return str;
    }
}
