/*
 * To change this license header, choose License Headers in Project Properties.
 * To change this template file, choose Tools | Templates
 * and open the template in the editor.
 */
package Loaders;

import Db.dbConn;
import SupplyChain.Manager;
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
public class load_subcounty extends HttpServlet {
HttpSession session;
String sub_county;
String query;
String where_clause;
    protected void processRequest(HttpServletRequest request, HttpServletResponse response)
            throws ServletException, IOException, SQLException {
        response.setContentType("Content-type: application/json");
        try (PrintWriter out = response.getWriter()) {
           session = request.getSession();
           dbConn conn = new dbConn();
           
          
          
         where_clause = "WHERE 1=1 AND"; 
          
         if(request.getParameter("county")!=null && !request.getParameter("county").equals("")){
           String[] counties = request.getParameter("county").split(",");  
          for(String county:counties){
              if(!county.equals("") && !county.equals(",")){
              county = county.replace("'", "\'");
              where_clause+=" county='"+county+"' OR ";
              }
          }
          }
          
          //remove last 4 characters
          where_clause = Manager.removeLastChars(where_clause, 3);
           
           
           
           
            JSONObject finalobj = new JSONObject();
            JSONArray jarray = new JSONArray();
            query = "";
            
              query = "SELECT DISTINCT(sub_county) AS sub_county FROM report "+where_clause+" ORDER BY sub_county";
              conn.pst = conn.conn.prepareStatement(query);
              conn.rs = conn.pst.executeQuery();  
            
            System.out.println("query : "+conn.pst);
            while(conn.rs.next()){
                JSONObject obj = new JSONObject();
                obj.put("sub_county", conn.rs.getString(1));
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
        Logger.getLogger(load_subcounty.class.getName()).log(Level.SEVERE, null, ex);
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
        Logger.getLogger(load_subcounty.class.getName()).log(Level.SEVERE, null, ex);
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
