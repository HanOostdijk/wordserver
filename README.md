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
    
Information about the Microsoft Word object model can be found e.g. in   
*   [*Object model (Word VBA reference)*](https://msdn.microsoft.com/en-us/library/office/ff837519.aspx)  
*   [*Word Enumerated Constants*](https://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx)  
    
Acknowledgement : copied most of the code from *wordreport* in this class:    
    (https://nl.mathworks.com/matlabcentral/fileexchange/17953-wordreport)
    Author: Laurent Vaylet
    E-mail: laurent.vaylet@gmail.com
    Release: 1.0
    Release date: 12/10/07
    Some extra functions were added to *wordreport* by Dmytro Makogon