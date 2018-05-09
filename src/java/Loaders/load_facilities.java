/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Loaders;

import Db.dbConn;
import java.io.IOException;
import java.io.PrintWriter;
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
public class load_facilities extends HttpServlet {
    HttpSession session;
    String county,sub_county,facility,mfl_code;
    String query;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
        response.setContentType("Content-type: application/json");
        try (PrintWriter out = response.getWriter()) {
           session = request.getSession();
           dbConn conn = new dbConn();
           
            JSONObject finalobj = new JSONObject();
            JSONArray jarray = new JSONArray();
            query = "";
            if(request.getParameter("sub_county")!=null && !request.getParameter("sub_county").equals("")){
              sub_county = request.getParameter("sub_county");
              query = "SELECT DISTINCT(mfl_code) AS mfl_code,facility AS health_facility FROM report WHERE sub_county =? ORDER BY health_facility";
              conn.pst = conn.conn.prepareStatement(query);
              conn.pst.setString(1, sub_county);
              conn.rs = conn.pst.executeQuery();
            }
            else if(request.getParameter("county")!=null && !request.getParameter("county").equals("")){
              county = request.getParameter("county");
              query = "SELECT DISTINCT(mfl_code) AS mfl_code,facility AS health_facility FROM report WHERE county =? ORDER BY health_facility";
              conn.pst = conn.conn.prepareStatement(query);
              conn.pst.setString(1, county);
              conn.rs = conn.pst.executeQuery();
            }
            else{
              query = "SELECT DISTINCT(mfl_code) AS mfl_code,facility AS health_facility FROM report ORDER BY health_facility";
              conn.pst = conn.conn.prepareStatement(query);
              conn.rs = conn.pst.executeQuery();  
            }
            System.out.println("query facility : "+conn.pst);
            while(conn.rs.next()){
                JSONObject obj = new JSONObject();
                obj.put("mfl_code", conn.rs.getString(1));
                obj.put("facility_name", conn.rs.getString(2));
                jarray.add(obj);
            }
            
            finalobj.put("data", jarray);
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
            Logger.getLogger(load_facilities.class.getName()).log(Level.SEVERE, null, ex);
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
            Logger.getLogger(load_facilities.class.getName()).log(Level.SEVERE, null, ex);
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
