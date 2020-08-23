<h2>Faster-ETL-When-You-Work-in-Excel</h2>
<h3>Intro</h3>
<p>In this project I take advantage of ADO which stands for ActiveX Data Obecjts. If you work with Excel and you need to get a data into your wokrsheet, ADO makes any type of data source accessible for you. It can be SQL Server, Access data base, or regular Excel worksheet.</p>
<p>A lot of work in excel depends on copy pasting data from excel to excel. Doing it manually, requires you to open each excel, copy data, and append the data to an exitistig one and close the source. We can simply automate this steps by writing a macro. However the mistake that people make is duplicating those steps in the code, having to open and close source excel. It is ok when you have 1, 2 or 3 sources excel. However when you work with multiple excels, this may be consuming a lot of time. Especially when you copy data by selecting range and making a copy method on it. That prolongs whole process as well.</p>
<p>Here come ADO to make it faset without opening each excel and without selecting range to copy in every source. With ADO you can just query Excel like any other SQL or Access database. That saves you a lot of time an ensures data integrity as there is no risk of selecting wrong range or too large or to small range of data to be copied.</p>
<h3>Early Binding</h3>
<ul>
  <li>There is need to activate the reference of Microsoft AcitveX Data Object X.X Library, where X.X is the most recent version you have.</li>
  <li><We need to get connection string:/li>
    <ul>
      <li>Provider=Microsoft.ACE.OLEDB.12.0;Data Source=c:\myFolder\myExcel2007file.xlsx;Extended Properties="Excel 12.0 Xml;HDR=YES";</li>
    </ul>
</ul>

<h3>Late Binding</h3>
<ul>
  <li>...</li>
  <li>...</li>
</ul>
