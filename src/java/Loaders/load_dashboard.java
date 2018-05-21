/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Loaders;

import Db.dbConn;
import java.io.IOException;
import java.io.PrintWriter;
import java.sql.ResultSetMetaData;
import java.sql.SQLException;
import java.util.logging.Level;
import java.util.logging.Logger;
import javax.servlet.ServletException;
import javax.servlet.http.HttpServlet;
import javax.servlet.http.HttpServletRequest;
import javax.servlet.http.HttpServletResponse;
import javax.servlet.http.HttpSession;
import org.json.simple.JSONArray;
import org.json.simple.JSONObject;

/**
 *
 * @author GNyabuto
 */
public class load_dashboard extends HttpServlet {
HttpSession session;
int sites_assessed,expected_sites;
String average_phamarcy,average_lab,average_score;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
        response.setContentType("text/html;charset=UTF-8");
        try (PrintWriter out = response.getWriter()) {
           session = request.getSession();
           dbConn conn = new dbConn();
           
            JSONObject finalobj = new JSONObject();
            JSONArray jarray = new JSONArray();
            
            JSONArray array_facilities = new JSONArray();
            JSONArray array_storage_score = new JSONArray();
            JSONArray array_inventory_score = new JSONArray();
            JSONArray array_RRM_score = new JSONArray();
            JSONArray array_MIS_score = new JSONArray();
            JSONArray array_EUV_lab_score = new JSONArray();
            JSONArray array_IM_additional_lab_score = new JSONArray();
            JSONArray array_MTC_pharm_score = new JSONArray();
            
            sites_assessed = 0;
            String getassessed = "SELECT count(id) FROM report";
            conn.rs = conn.st.executeQuery(getassessed);
            if(conn.rs.next()){
                sites_assessed = conn.rs.getInt(1);
            }
            String getexpected = "SELECT number_expected FROM expected_sites WHERE id=1";
            conn.rs = conn.st.executeQuery(getexpected);
            if(conn.rs.next()){
                expected_sites = conn.rs.getInt(1);
            }
            finalobj.put("sites_assessed", sites_assessed);
            finalobj.put("expected_sites", expected_sites);
            
            //GET AGGREGATE SCORE
            
            String getscore = "SELECT county AS 'County'," +
                "ROUND(AVG(IFNULL(storage_total_score*100,0))) AS 'Storage areas score',ROUND(AVG(IFNULL(inventory_total_score*100,0))) AS 'Inventory management Score' ," +
                "ROUND(AVG(IFNULL(RRM_total_score*100,0))) AS 'Resources and reference materials score', ROUND(AVG(IFNULL(MIS_total_score*100,0))) AS 'Availability and use of MIS Tools Score', " +
                "ROUND(AVG(IFNULL(EUV_lab_score*100,0))) AS 'Laboratory HIV RTKs End Use Verification score', " +
                "ROUND(AVG(IFNULL(IM_additional_lab_score*100,0))) AS 'Laboratory Inventory Management: additional commodities score'," +
                "ROUND(AVG(IFNULL(MTC_pharm_score*100,0))) AS  'Pharmacy Medicines and Therapeutics Committees Score' FROM report " +
                "GROUP BY County";
            conn.rs = conn.st.executeQuery(getscore);
            
            ResultSetMetaData metaData = conn.rs.getMetaData();
            int count = metaData.getColumnCount(); //number of column 
               String header = metaData.getColumnLabel(1);
               
            while(conn.rs.next()){
            array_facilities.add(conn.rs.getString(1));
            array_storage_score.add(conn.rs.getString(2));
            array_inventory_score.add(conn.rs.getString(3));
            array_RRM_score.add(conn.rs.getString(4));
            array_MIS_score.add(conn.rs.getString(5));
            array_EUV_lab_score.add(conn.rs.getString(6));
            array_IM_additional_lab_score.add(conn.rs.getString(7));
            array_MTC_pharm_score.add(conn.rs.getString(8));
           
            }
            JSONArray all_datasets = new JSONArray();
            
            JSONObject obj_storage = new JSONObject();
            obj_storage.put("label", "Storage Areas");
            obj_storage.put("data", array_storage_score);
            obj_storage.put("borderColor", "rgba(0, 123, 255, 0.9)");
            obj_storage.put("borderWidth", 0);
            obj_storage.put("backgroundColor", "#6E3505");
            
            all_datasets.add(obj_storage);
            
            
            
            JSONObject obj_inventory = new JSONObject();
            obj_inventory.put("label", "Inventory Management");
            obj_inventory.put("data", array_inventory_score);
            obj_inventory.put("borderColor", "rgba(0, 123, 255, 0.9)");
            obj_inventory.put("borderWidth", 0);
            obj_inventory.put("backgroundColor", "#B8B204");
            
            all_datasets.add(obj_inventory);
            
            
            
            JSONObject obj_RRM = new JSONObject();
            obj_RRM.put("label", "Resources and Reference Materials");
            obj_RRM.put("data", array_RRM_score);
            obj_RRM.put("borderColor", "rgba(0, 123, 255, 0.9)");
            obj_RRM.put("borderWidth", 0);
            obj_RRM.put("backgroundColor", "#04B875");
            
            all_datasets.add(obj_RRM);
            
            
            
            JSONObject obj_MIS = new JSONObject();
            obj_MIS.put("label", "Availability and Use of MIS Tools");
            obj_MIS.put("data", array_MIS_score);
            obj_MIS.put("borderColor", "rgba(0, 123, 255, 0.9)");
            obj_MIS.put("borderWidth", 0);
            obj_MIS.put("backgroundColor", "#7F04B8");
            
            all_datasets.add(obj_MIS);
            
            
            
            JSONObject obj_EUV = new JSONObject();
            obj_EUV.put("label", "HIV RTKs End Use Verification");
            obj_EUV.put("data", array_EUV_lab_score);
            obj_EUV.put("borderColor", "rgba(0, 123, 255, 0.9)");
            obj_EUV.put("borderWidth", 0);
            obj_EUV.put("backgroundColor", "#4104B8");
            
            all_datasets.add(obj_EUV);
            
            
            
            JSONObject obj_IM = new JSONObject();
            obj_IM.put("label", "Inventory Management: Additional Commodities");
            obj_IM.put("data", array_IM_additional_lab_score);
            obj_IM.put("borderColor", "rgba(0, 123, 255, 0.9)");
            obj_IM.put("borderWidth", 0);
            obj_IM.put("backgroundColor", "#042FB8");
            
            all_datasets.add(obj_IM);
            
            
            
            JSONObject obj_MTC = new JSONObject();
            obj_MTC.put("label", "Medicines and Therapeutics Committees");
            obj_MTC.put("data", array_MTC_pharm_score);
            obj_MTC.put("borderColor", "rgba(0, 123, 255, 0.9)");
            obj_MTC.put("borderWidth", 0);
            obj_MTC.put("backgroundColor", "#33FFF7");
            
            all_datasets.add(obj_MTC);
            
            JSONObject obj_data = new JSONObject();
            obj_data.put("labels", array_facilities);
            obj_data.put("datasets", all_datasets);
            
            String average_scorequery = "SELECT AVG(total_pharm_score)*100 AS total_pharm_score,AVG(total_lab_score)*100 AS total_lab_score,AVG(total_score)*100 AS total_score FROM report";
            conn.rs = conn.st.executeQuery(average_scorequery);
            if(conn.rs.next()){
              average_phamarcy = conn.rs.getString(1);
              average_lab = conn.rs.getString(2);
              average_score = conn.rs.getString(3); 
            }
            finalobj.put("average_phamarcy", average_phamarcy);
            finalobj.put("average_lab", average_lab);
            finalobj.put("average_score", average_score);
            
            finalobj.put("overall_score", obj_data);
            
            out.println(finalobj);
        }
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
        Logger.getLogger(load_dashboard.class.getName()).log(Level.SEVERE, null, ex);
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
        Logger.getLogger(load_dashboard.class.getName()).log(Level.SEVERE, null, ex);
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

}
