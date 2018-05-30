# GnuPlotForExcel

Welcome to the GnuPlot add-in for Excel! 

For Excel 2010, 2013, and 2016. 

This add-in is alternative to Excel graphics which facilitates use of GnuPlot directly in Excel sheets. GnuPlot scripts are stored in Excel along with the generated graphs and can be instantly edited and regenerated without leaving the Excel environment. Data can be read from Excel sheets by a cell range ("A1:C400") or using an Excel named range. 

Step 0 - Install GnuPlot if you haven't already. 

https://sourceforge.net/projects/gnuplot/ 

Note the installation directory of the executable, typically C:\Program Files\gnuplot\bin, you will need it later. 

Step 1 - Download the GnuPlot add-in for Excel from GitHub.

Minimally, you will need only GnuPlot.xlam. GnuPlot.xlsx is an empty workbook with GnuPlot buttons in the context menus.  GnuPlot.xltx is an Excel template which defines the relevant buttons. See the following for more technical information on context menus. 

https://gregmaxey.com/word_tip_pages/customize_shortcut_menu.html

https://www.rondebruin.nl/win/s2/win014.htm 

Step 2 - Make sure the downloaded files are unblocked. 

When add-in files are downloaded from the Internet, they are blocked by Windows for "security" reasons. Be sure to unblock GnuPlot.xlam, GnuPlot.xlsx, and GnuPlot.xltx immediately after downloading them. 

https://blogs.msdn.microsoft.com/delay/p/unblockingdownloadedfile/ 

https://winaero.com/blog/disable-downloaded-files-from-being-blocked-in-windows-10/ 

http://www.jkp-ads.com/Articles/Excel-Add-ins-fail-to-load.asp 

Step 3 - Install the add-in. 

The addin needs to be installed and activated in Excel. This procedure varies slightly depending on your versions of Windows and Excel.

http://www.contextures.com/exceladdins.html 

https://www.excelcampus.com/tools/how-to-install-an-excel-add-in-guide/ 

https://www.youtube.com/watch?v=reuU2zUsEPM 

Hard part's over.

Step 4 - Verify that the add-in has been successfully installed. 

You should see a new tab in the ribbon, GNUPLOT, toward the right. When you click on it, you should see icons for Setup, Create, Edit, and Render, along with Help and Info. When you click on Info, you should see a simple info box. If this works, the add-in has been successfully installed.

Step 5 - Set up the add-in for first time use. 

The add-in needs to know where the GnuPlot executable can be found on your computer. Often it is C:\Program Files\gnuplot\bin, but it can be anywhere. Click the Setup button in the GNUPLOT ribbon and fill in the relevant field in the dialogue box. This information will be saved in your computer's registry and needs to be entered only once. 

Step 6 - Use the GnuPlot.xlsx helper workbook and/or the GnuPlot.xltx template. 

This is an empty workbook with the context menus configured so that you don't need to click around the ribbon to find the options, instead they are right there when you right click.  Alternatively, the GnuPlot.xltx is a normal Excel template which accomplishes the same thing.  Note there is no VBA code here, only ribbon XML definitions.

Step 7 - Your first GnuPlot 

As a first example, you will get one of the existing demos to work. This demo generates its own data and does not require data from outside the graph. The catalogue of demos is here http://gnuplot.sourceforge.net/demo_5.0/ Start Excel with a fresh new workbook. Click the Create button on the ribbon. A square Gnuplot logo should appear. Stretch the logo to be the approximate shape desired. Click the Edit button on the ribbon. An editor pane should pop up. Copy the entire GnuPlot script from http://gnuplot.sourceforge.net/demo_5.0/contours.6.gnu and paste it into the editor window, make no changes, and press OK. Click the Render button on the ribbon. The demo plot should appear. You may want to make some adjustments to the size of the graph, then press Render again. The resulting picture can be copied freely to other Excel sheets, to Word, or to PowerPoint. 

Step 8 - Your second GnuPlot 

Next, you will get another existing demo to work, but one of them which requires data. In your workbook, create a sheet called "Finance". In your GnuPlot installation locate file "\gnuplot\demo\finance.dat" and open it with Notepad or equivalent. Copy the entire contents of the file to the clipboard, Ctrl-A then Ctrl-C. Go to your newly created Finance sheet and paste the clipboard contents starting at the first cell. Ctrl-V to paste. Make sure the entire contents of what you just pasted is selected. Type the word Finance and press enter in the name box of the spreadsheet, which should be directly above cell A1. This will establish a Named Range called Finance. https://exceljet.net/named-ranges for more info. Create a new sheet or use one of the existing empty ones. Click the Create button, resize the picture, then click the Edit button, as before. Copy the entire GnuPlot script from http://gnuplot.sourceforge.net/demo_5.2/finance.4.gnu In the editor window, change 'finance.dat' to <<"finance!finance">> then press save. The first "finance" is the name of the sheet, while the second "finance" is the named range. Click the Render button on the ribbon. The demo plot should appear. As before, You may want to make some adjustments to the size of the graph, then press Render again. 

That's it.
