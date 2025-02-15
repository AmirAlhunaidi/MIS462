<%@ Language=VBScript %>
<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">
<html xmlns="http://www.w3.org/1999/xhtml">
<!-- DW6 -->
<!-- #include virtual ="/Code/constants.asp" -->
<!-- #include virtual ="/BUSA331/ClassConstants.asp" -->

<html>
<head>
<title>MIS462Summative03</title>
<style type="text/css">
<!--
.style2 {
	color: #0000FF
}
-->
</style>
<meta http-equiv="Content-Type" content="text/html; charset=iso-8859-1" />
</head>

<body bgcolor="#FFFFFF" text="#000000">
<form action="https://eprofessor.azurewebsites.net/code/ProcessFormData.asp" method="post" name="frmHW" id="frmHW">
  <h2 align="center"><i><font color="#cc00FF">MIS 462
    <% response.write(semester) %>
    Summative Assignment</font></i></h2>
  <table width="100%" border="1" cellspacing="1" cellpadding="1">
    <tr bgcolor="#00FFFF">
      <td >Email Address:
        <input type="text" name="email" id="email" value="ValidEmail@winona.edu" /></td>
      <td >First Name:
        <input type="text" name="FirstName" id="FirstName" size="30" maxlength="50" /></td>
      <td >Last Name:
        <input type="text" name="LastName" id="LastName" size="25" maxlength="50" /> </td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td>Semester:
        <input type="text" name="Semester" id="Semester" value=<% response.write(semester)%> />      
      </td>
      <td>Class:
        <input type="text" id="Class" name="Class" Value="MIS462"/></td>
      <td>StarID:
        <input type="text" name="PIN" /><input name="InstID" type="hidden" id="InstID" value="00617282" /></td>
    </tr>
    <tr bgcolor="#00FFFF">
      <td >Section:
        <input name="Section" id="Section" value="01"/>
       </td>
      <td > Assignment:
        <input name="Assignment" id="Assignment" value="Summative03"/>
      </td>   
      <td>&nbsp;</td>
      
      
      
    </tr>
    <tr bgcolor="#00FFFF">
      <td colspan="3">&nbsp;</td>
    </tr>
    <tr bgcolor="#FF0000">
      <td colspan="3"><div align="center">
          <input type="submit" name="Submit" value="Submit" />
        </div></td>
    </tr>
  </table>
  <p><br>
    <font color="#FF0000"><strong>1000 
      points</strong></font></p>
  <h2  align="center"><font color="#cc00FF"><em><u> Problems, Chapters  19 to 26 and 87 </u></em></font></h2>
  <p  align="left"><br>
    Do all of these problems in a single Excel Workbook.<br>
    Solve each problem on a separate Worksheet.<br>
    Your Excel File must be setup like you did in Formative01 with appropriate Worksheet tab names.
  </p>
  <p align="left">
    <br>
  <strong>(200) </strong> When done, upload your file, named Summative03.xlsx, to the D2L Assignment Folder 'Summative03'.</p>
  <hr>
  <hr>




  
  <p align="left">Counting Cells-Chapter 19 Problem 7</p>
<p align="left">Refer to the file Rock.xlsx </p>
<div class="Section115">
  <p><span class="style2"><strong>(50) 1. For the cell range D4:G15, count the cells containing a numeric value.</strong></span><br />
    <span class="normal"><font color="#0000FF">
      <textarea name="q1" cols="120" rows="3">Number numeric cells:</textarea>
      </font></span><br />
    </p>
  <p><span class="style2"><strong>(50) 2. Count the number of blank cells.</strong></span><br />
    <span class="normal"><font color="#0000FF">
      <textarea name="q2" cols="120" rows="3" id="q2">Number of blank cells:</textarea>
    </font></span>    </p>
  <p><br />
    <span class="style2"><strong>(50) 3. Count the number of nonblank cells.</strong></span><br />
    <o:p><span class="normal"><font color="#0000FF">
      <textarea name="q3" cols="120" rows="3" id="q3">Number of non blank cells:</textarea>
  </font></span></o:p></p>
</div>
<div class="Section1"></div>
<hr />
<p >
  <o:p>Summing Cells-Chapter 20 Problem 2 </o:p></p>

  <p>Use the file Makeup.xlsx to answer this question:<br />
  <div class="style2">
    <strong>(50) 4. Use the SUMIF function to determine the total revenue earned before December 10, 2005</strong> </p>
  <p class="normal"><span class="normal"><font color="#0000FF">
    <textarea name="q4" cols="120" rows="3" id="q4">Total Revenue:</textarea>
    </font></span></p>
</div>
<hr />
<p >
  <o:p>Offset-Chapter 21 Problem 1</o:p></p>
<p  ><span class="style2"><strong>5. (<strong>50</strong>) </span></span></o:p><o:p> </o:p><o:p></o:p></span><span class="style2 style2"><o:p></o:p></span><span class="style2"><o:p></o:p><strong>
<o:p>The file C21p1.xlsx supplies data about the units sold for 11 products during the years 1999&ndash;2003.<br />
  Write a formula using the MATCH (refer to Chapter 5) and OFFSET functions that picks up the sales of a given product during a given year.</o:p></strong></span><span class="style6"><o:p></o:p></span><span class="style2"><o:p><br />
  <span class="normal"><font color="#0000FF">
  <textarea name="q5" cols="120" rows="3" id="q5">Formula:</textarea>
  </font></span><br />
</o:p></span></p>
<div class="Section19">
  <hr />
</div>
<p >Indirect-Chapter 22 Problem 2</p>
<div class="Section117">
  <p class="style2"><strong>6. (50) The workbook P22_2.xlsx contains data for the sales of five products in four regions (East, West, North, and South).<br />
    Use the INDIRECT function to create formulas that enable you to easily add up the total sales of any combination of consecutively numbered products, such as Products 1&ndash;3, Products 2&ndash;5, and the like.<br />
    Paste the forumla below:
    </strong></p>
</div>
<div class="Section110">
  <p class="normal"><font color="#0000FF">
    <textarea name="q6" cols="120" rows="3" id="q6">Indirect formula:
</textarea>
    <br />
    </font></p>
</div>
<div class="Section1">
  <hr />
  <p class="normal">Conditional Formatting-Chapter 23 Problem 1</p>
  <p class="normal"> Using the data in the file SandP.xlsx, use conditional formatting in the following situations:</p>
  <div class="Section118">
    <p class="normal">Format in bold each month in which the value of the S&amp;P increased and underline each month in which the value of the S&amp;P decreased.</p>
    <p class="normal">Highlight in green each month in which the S&amp;P changed by at most 2 percent.</p>
    <p class="normal">Highlight the largest S&amp;P value in red and the smallest value in purple.</p>
    </div>
  <div class="Section111"> </div>
  <div class="Section15">
    <table   border="2" cellpadding="2" cellspacing="2">
      <tr>
        <td  valign="top">(<strong>50</strong>) 7. Make a screen shot of Conditional Formating Rules Manager dialog box for the S&amp;P values, save it in a convenient location as you will combine it into one pdf file when you complete this assignment. <br />
          <a href="../../Formative/Formative00/Formative00.asp">Refer to Assignment Formative00, Exercise 7</a> for details.</td>
        </tr>
      </table>
    <hr />
  </div>
</div>
<div class="Section11">
  <h1 class="normal">Conditional Formatting-Chapter 23 Problems 11, 12, 13</h1>

  <p class="normal">The file Nbasalaries.xlsx contains salaries of NBA players in millions of dollars. <br />
    Set up    data bars to summarize this data.<br />
Players making less than $1 million should have the    shortest data bar, and players making more than $15 million should have largest data    bar.<br />
  </p>
  <table   border="2" cellpadding="2" cellspacing="2">
    <tr>
      <td  valign="top"> (<strong>50</strong>) 8. Make a screen shot of Conditional Formating showing the data bars, save it in a convenient location as you will combine it into one pdf file when you complete this assignment. <br />
          <a href="../../Formative/Formative00/Formative00.asp">Refer to Assignment Formative00, Exercise 7</a> for details.</td>
    </tr>
  </table>
  <hr />
  <p class="normal"> Set up a three-color scale to summarize the NBA salary data. <br />
    Change the color of the    bottom 10 percent of all salaries to green and the top 10 percent to red.<br />
  </p>
  <table   border="2" cellpadding="2" cellspacing="2">
    <tr>
      <td  valign="top"> (<strong>50</strong>) 9. Make a screen shot of Conditional Formating showing the three color scale, save it in a convenient location as you will combine it into one pdf file when you complete this assignment. <br />
          <a href="../../Formative/Formative00/Formative00.asp">Refer to Assignment Formative00, Exercise 7</a> for details.</td>
    </tr>
  </table>
  <hr />
  <p class="normal">Using the data in the file NBASalaries.xlsx, use five icons to summarize the NBA Player salary data.<br />
    Create break points at $3 million, $6 million, $9 million and $12 million.
    <br />
    Use five colored arrows for the icon sets.
  </p>
  <div class="Section111"></div>
  <div class="Section15">
    <table   border="2" cellpadding="2" cellspacing="2">
      <tr>
        <td  valign="top"> (<strong>50</strong>) 10. Make a screen shot of Conditional Formating showing the first few cells, save it in a convenient location as you will combine it into one pdf file when you complete this assignment. <br />
          <a href="../../Formative/Formative00/Formative00.asp">Refer to Assignment Formative00, Exercise 7</a> for details. </td>
      </tr>
    </table>
    <hr />
  </div>
  <div class="Section1">
    <div class="Section15">
      <hr />
    </div>
  </div>
  <div class="Section11">
    <p class="normal">Tables-Chapter 25 Problem 7</p>
    <p class="normal">(<strong>50</strong>) Use the data in the file NikeData.xlsx, which contains quarterly sales revenues for Nike.<br />
      Create a table, named 'NikeData'<br />
      Choose an appropriate format.<br />
      Create a chart, 2-d columnar.</p>
    <p class="normal">Add the following data:<br />
    </p>
    <table width="45%" border="1">
      <tr>
        <td>Year/quarter</td>
        <td>Revenue</td>
      </tr>
      <tr>
        <td>2001 1</td>
        <td>3000</td>
      </tr>
      <tr>
        <td>2001 2</td>
        <td>3100</td>
      </tr>
      <tr>
        <td>2001 3</td>
        <td>3200</td>
      </tr>
      <tr>
        <td>2001 4</td>
        <td>3800</td>
      </tr>
    </table>
    <br />
    <div class="Section15">
      <table  border="2" cellpadding="2" cellspacing="2">
        <tr>
          <td valign="top"> (<strong>50</strong>) 11. Make a screen shot of Conditional Formating showing the table and the 2-d columnar graph depicting the new data entered, save it in a convenient location as you will combine it into one pdf file when you complete this assignment. <br />
          <a href="../../Formative/Formative00/Formative00.asp">Refer to Assignment Formative00, Exercise 7</a> for details.</td>
          </tr>
        </table>
      <hr />
    </div>
  </div>
  <p class="normal">Controls-Chapter 26 Problem 1, spin buttons</p>
  <p class="normal"> Add a spin button to the car NPV example (NPVspinners.xlsx) that allows the tax rate to vary between 30 and 50 per cent.</p>
  <table   border="2" cellpadding="2" cellspacing="2">
    <tr>
      <td  valign="top"> (<strong>50</strong>) 12. Make a screen shot of the tax rate spinner control set to 34 percent, save it in a convenient location as you will combine it into one pdf file when you complete this assignment. <br />
          <a href="../../Formative/Formative00/Formative00.asp">Refer to Assignment Formative00, Exercise 7</a> for details.</td>
    </tr>
  </table>
  <hr />
  <p> <h2>Chapter 87-Array Formulas </h2>
  
  (<strong>100</strong>) 13. Use the workbook <a href="ArrayProblem.xlsx">ArrayProblem.xlsx</a><br />
  <p class="style2"> Enter an array formula in cell D11 that will calculate the grand total wages owed.
Do not create any intermediate formulas.</p>
<input type="text" name="q13" id="q13" size="50" maxlength="50" value="Grand Total Wages: " />
</p>
  
  
  
  <hr />
  <p> (<strong>100</strong>) 14. Combine your six screen shots from the previous exercises in order into one .pdf file named Summative03ScreenShots.pdf and upload this file to the D2L Dropbox folder 'Summative03 Screen Shots' </p>


  <hr>
  <hr>
  
   <table width="100%" border="1" cellpadding="1" cellspacing="1">
   <tr bgcolor="#FF0000">
     <td ><div align="center">
           <input type="submit" name="Submit2" value="Submit" />
         </div>
     </td>
   </tr>
   </table>

  
</form>
</body>
</html>
