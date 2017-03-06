# wordserver

**A MATLAB activex class to access Microsoft Word documents**

It is an extension of *wordreport* by Laurent Vaylet.

Main changes:   
* converted code to a class
* included/expanded inclusion of figures and tables with captions and bookmarks
* included references to figures and tables
* included header and footer
* changed handling of arguments
* included all wd constants (that are used) in one function (with translate function)

Copyright 2017 Han Oostdijk  MIT License  
Version: 1.0  Date 06MAR2017
    
Acknowledgement : copied most of the code from [*wordreport*] (https://nl.mathworks.com/matlabcentral/fileexchange/17953-wordreport) by Laurent Vaylet (E-mail: laurent.vaylet@gmail.com) in this class. In that version (Release 1.0 of 12DEC2007) some extra functions were added by Dmytro Makogon.

Suggestion for changing or expanding this class:  
A way to add functionality to this class is by using the macro recorder facility in Microsoft Word. By editing the macro (e.g. replacing a sequence of operations with a loop) a VBA program with the required functionality is created. It is then relatively easy to convert this program (manually) to a MATLAB function. The functions that are already available in the **wordserver** class can serve as examples. Information about the Microsoft Word object model can be found e.g. in   
*   [*Object model (Word VBA reference)*](https://msdn.microsoft.com/en-us/library/office/ff837519.aspx)  
*   [*Word Enumerated Constants*](https://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx)  
    
This repository also contains an example **wordserver_example.m** that shows how to insert a figure and a table with the corresponding references. 
