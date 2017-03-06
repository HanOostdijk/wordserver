classdef wordserver  < handle
%{
Class to access Microsoft Word documents.
This is an extension of 'wordreport' by Laurent Vaylet 
Main changes:
    converted code to a class
    included/expanded inclusion of figures and tables with captions and bookmarks
    included references to figures and tables
    included header and footer
    changed handling of arguments
    included all wd constants (that are used) in one function (with translate function)
Copyright 2017 Han Oostdijk  MIT License 
Version: 1.0  Date 10feb2017
    
Information about the Microsoft Word object model can be found e.g. in 
    https://msdn.microsoft.com/en-us/library/office/ff837519.aspx
    https://msdn.microsoft.com/en-us/library/office/aa211923(v=office.11).aspx
    
Acknowledgement : copied all code from 'wordreport' in this class:    
    https://nl.mathworks.com/matlabcentral/fileexchange/17953-wordreport
    Author: Laurent Vaylet
    E-mail: laurent.vaylet@gmail.com
    Release: 1.0
    Release date: 12/10/07
    Some extra functions were added to 'wordreport' by Dmytro Makogon
    
Todo :
    SetTableColWidth        : allow specification of specific columns 
    SetTableColAlign        : allow specification of specific columns 
    
    All functions with exception of close_doc, saveas_doc, save_doc, activate_doc,
    set_saved_doc and SelectionSetFont have a structure od with default parameters.
    These parameters can be overwritten in three ways:
        with a structure optionally followed by overwriting name-value pairs 
        with a cell array optionally followed by overwriting name-value pairs 
        with name-value pairs
    
    functions :
        wordserver              : Starts word activex server and sets default style and prefix for headers
            parameters: 
                'normal' (default 'normal') 
                'heading' (default 'Heading')
        delete                  : Destructor for class
            no parameters
        quit                    : Quit the Word application
            no parameters
        set_visible             : make Word (in-)visible
            parameters: 
                'visible' (default true) 
        open_doc                : Open a new or existing word document    
            parameters: 
                'file_name' (default '') 
        close_doc               : Close document without saving first  
            parameters: 
                the document object that is to be closed 
        saveas_doc              : Save document (with another name?)
            parameters: 
                the document object that is to be closed
                'file_name' (default is the old file name) 
        save_doc                : Save document (under its old name)
             parameters: 
                the document object that is to be closed    
        activate_doc            : Make a document the active one
             parameters: 
                the document object that is to be closed    
        get_active_doc          : Retrieve the active document object
             no parameters 
        set_saved_doc           : Indicate that the document does not need to be saved
             no parameters 
        AddParagraph            : Add a paragraph at the end of the document
            parameters: 
                'color' (default 'wdAuto') : color of paragraph as one of enum WdColorIndex
        AddSymbol               : Add a symbol represented by an integer
            parameters: 
                'symbol' (default 176) : an integer found in the Insert/Symbol menu
        SelectionSetFont        : Set font attributes for selection
            parameters: 
                one or more of the following fields without defaults but given with an example
                (only in struct of name-value pair format)
                'Name'                  'Arial'  
            	'Size'                  9    
            	'Bold'                  false  
            	'Italic'                false  
            	'Underline'             false 
            	'StrikeThrough'         false 
            	'ColorIndex'            'wdAuto'  
            	'DoubleStrikeThrough'   false 
            	'Superscript'           false   
            	'Subscript'             false   
        SelectionSetAlignment   : Set alignment for selection
            parameters: 
                'align' (default wdAlignParagraphCenter) : alignment one of enum WdParagraphAlignment
        SelectionInsertText     : Insert text over or besides selection
            parameters: 
                'text'          (default '')            : text to insert
                'lineBreaks'    (default [0,0])         : vector with line breaks before and after text
                'color'         (default 'wdAuto')      : color to use as one of enum WdColorIndex
                'fun'           (default 'InsertAfter') : function to use (TypeText, InsertBefore or InsertAfter )            
                'collapse'      (default true)         	: Collapse section after insert: true (yes: at end ), false (no)
        InsertTable             : Insert a table in the document with caption and bookmark
            parameters: 
                'dataCell'    	(default [])            : (character) data to insert 
                'lineBreaks'    (default [1,1])         : vector with line breaks before and after table
                'bmn',          (default '')            : bookmark name to associate with table
              	'captitle'      (default '')         	: caption to be given to table
                'cappos'  (default 'wdCaptionPositionBelow') : position of caption: one of enum WdCaptionPosition
                'align'   (default 'wdAlignRowCenter') 	: alignment one of enum WdRowAlignment
                'title'         (default '')            : title of table (in table properties)                   
                'descr'         (default '')            : description of table (in table properties) 
        InsertFigure            : Insert a figure in the document
            parameters: 
                'what'          (default [])            : graphics handle or graphics file name 
                'lineBreaks'    (default [1,1])         : vector with line breaks before and after figure
                'bmn',          (default '')            : bookmark name to associate with figure
              	'captitle'      (default '')         	: caption to be given to figure
                'cappos'  (default 'wdCaptionPositionBelow') : position of caption: one of enum WdCaptionPosition
                'align'   (default 'wdAlignRowCenter') 	: alignment one of enum WdRowAlignment
                'title'         (default '')            : title of figure (in figure properties)                   
                'descr'         (default '')            : description of figure (in figure properties) 
        InsertXRefCaption       : Insert cross reference to a caption
            parameters: 
                'bmn',          (default '')            : bookmark name associated with table, figure, ...
              	'reftype'  (default 'wdCaptionTable')   : ReferenceType one of enum WdReferenceKind 
                'refkind'  (default 'wdOnlyLabelAndNumber') : ReferenceKind one of enum WdReferenceKind
                'ashyper'       (default true)          : Insert as hyperlink
        SetTableColWidth        : Set column widths for table
            parameters:     
                'columns'      	(default [])            : column numbers to change ([] is all)
                'widths'      	(default [])            : widths in cm or as a percentage
              	'width_type'  	(default 'wdPreferredWidthPoints'): width type one of enum WdPreferredWidthType    
        SetTableColAlign        : Set alignment for columns for table
            parameters: 
                'columns'      	(default [])            : column numbers to change ([] is all)
                'align'      	(default [])            : cell array with WdParagraphAlignment constants
        SelectTableColumn       : Select a column of a table
            parameters: 
                'colnr'      	(default 1)             : number of column to select
        SelectTableRow          : Select a row of a table
            parameters: 
                'rownr'      	(default 1)             : number of row to select
        SelectTableCell         : Select a cell of a table
            parameters: 
                'rownr'      	(default 1)             : row number of the cell to select
                'colnr'      	(default 1)             : column number of the cell to select
        GetTable                : Get data from current table
            parameters: 
                'ignoreHeader' 	(default true)       	: ignore row header ?
        WriteTable              : Write data to current table
            parameters: 
                'data'          (default [])          	: (character) data to write to table
                'fstRow'     	(default 1)          	: number of the first row to write to
                'fstCol'       	(default 1)          	: number of the first column to write to
        SetRowNumb              : Change number of rows in current table
            parameters: 
                'nbRows'      	(default 1)          	: set number of rows
        SetColumnNumb           : Change number of cols in current table
            parameters: 
                'nbCols'      	(default 1)          	: set number of columns
        SetStyle                : Set current text style, used later by SelectionInsertText
            parameters: 
                'style'      	(default obj.normal)  	: set style to use
        SelectionInsertBreak    : Insert a break of specified type at start of selection 
            parameters: 
                'breaktype'   	(default 'wdPageBreak') : one of enum WdBreakType  
        Goto                    : Jump to specified location in document
            parameters: 
                'what'   	(default 'wdGotoBookmark')  : one of enum enum WdGoToItem   
                'which'    	(default 'wdGotoAbsolute')  : one of enum enum WdGoToDirection   
                'count'   	(default 1)                 : number of what  
                'name'    	(default '')                : name of what
                'delete' 	(default false)             : delete contents indicated selection  
        FindText                : Find text in document
            parameters: 
                'textToFind'   	(default '')            : text to find
                'forward'    	(default true)          : forward (true) or backward (false)    
                'replacement'  	(default '')            : number of what  
                'name'          (default '')           	: replacement text
                'findopts'      (default struct())   	: find options that overwrite the default find options:    
                    'wdReplace'         (default 'wdReplaceNone') 	:
                    'wdFindwrap'        (default 'wdFindContinue') 	:
                    'MatchCase'         (default false)             :
                    'MatchWholeWord'    (default false)             :
                    'MatchWildcards'    (default false)             :
                    'MatchSoundsLike'   (default false)             :
                    'MatchAllWordForms' (default false)             :
                    'Format'            (default false)             :
        Select                  : Extend or move selection 'right', 'left', 'home' or 'end' , 'up' or 'down' 
            parameters: 
                'direction'   	(default 'right')    	: direction ('right', 'left', 'home' or 'end' , 'up', 'down' 
                'unit'          (default 'wdCharacter')	: unit (one of enum WdUnits)   
                'count'         (default 1)             : number of 'units' 
                'extend'     	(default 'wdMove')   	: 'wdMove' or 'wdExtend'
        CreateTOC               : Create the table of contents, list of figures or list of tables    
            parameters: 
                'type'          (default 'TOC')     	: type TOC, Figure or Table
                'text' 	(default 'Table of Contents)   	: text to print above TOC or list
                'fontstruct'	(default  [])           : struct for use in SelectionSetFont to print text
                'tableader' (default 'wdTabLeaderDots')	: enum WdTabLeader
                'include_label'	(default  true)         : include label in list of figure or table
             	'UseHeadingStyles'	(default true)   	: forced to false for Figure and Table         
             	'UpperHeadingLevel'	(default  1)        : highest level to include in TOC      
             	'LowerHeadingLevel'	(default  3)        : lowest level to include in TOC        
               	'UseFields'         (default  false)   	:             
              	'TableID'           (default  '')   	:                   
               	'RightAlignPageNumbers'	(default  true)	:      	 
              	'IncludePageNumbers'	(default  true)	:        
               	'AddedStyles'       (default '')        :                 
               	'UseHyperlinks'     (default true)   	:          	
               	'HidePageNumbersInWeb'	(default true)	:   
               	'UseOutlineLevels'	(default false)   	:     
        UpdateTOC               : Update the Table of Contents 
            parameters: 
                'upd_pn_only'       (default false) 	: update page numbers only ?
        AddHeaderFooter         : Add header or footer to the document
            parameters: 
                'align'         (default 'wdCenter') 	: alignment one of WdAlignmentTabAlignment
                'infooter'   	(default true)          : true when in footer otherwise false
                'pagetxt'    	(default 'page ')       : prefix for page number
                'inc_page'   	(default true)          : true when page number is to be included
                'inc_tot'       (default true)          : true when total number of pages isto be included
                'inctxt'        (default ' of ')        : prefix for total number
        PrintMethods            : Print all available methods for a Word ActiveX object    
            parameters: 
                'category'     	(default 'Application')         : object type for which methods are to be listed
                'headingString'	(default [obj.heading,' 2'])  	: heading style for this list
%}
   
    properties
        w                                               % handle to activex word server
        currentStyle                                    % style to use for next action
        wd_alpha_list                                	% word constants (alpha)
        wd_num_list                                 	% word constants (numeric)
        normal                                          % default style (in English Normal)
        heading                                         % default header prefix (in English Heading)
    end
    methods
        function obj = wordserver(varargin)             % constructor for class
            % WORDSERVER Starts word activex server and sets default style and prefix for headers
            od = struct( ...                            % default constants
                'normal', 'Normal', ...                 % for default style
                'heading', 'Heading' ) ;                % for prefix of header
            obj.w = wordserver. ...                     % try to reuse existing server
                actxserver2('Word.Application');
            [obj.wd_alpha_list,obj.wd_num_list] = ...   % convert the array with 
                wordserver.wd_def_const() ;             %  word constants
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults              
            obj.normal          = od.normal ;                   % set the name of the default style           
            obj.heading         = od.heading ;                  % set the prefix for headings          
            obj.currentStyle    = obj.normal;          	% set current style to the default style
        end
        function delete(obj)                          
            % DELETE Destructor for wordserver class
            try
                quit(obj)                               
            catch
            end
        end
        function quit(obj)
            % QUIT Quit the Word application
            invoke(obj.w, 'Quit');                      % quit the Word application
            delete(obj.w);                              % free storage related to server
        end
        function set_visible(obj,varargin)
            %SET_VISIBLE make Word (in-)visible
            od = struct(    ...                         % defaults for arguments
                'visible', true  ...                    % make Word (in-)visible
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults   
            set(obj.w,'Visible',od.visible);            % make Word visible or not
        end
        function doc = open_doc(obj,varargin)
            %OPEN_DOC Open a new or existing word document
            od = struct(    ...                         % defaults for arguments
                'file_name', ''  ...                    % file name
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults  
            h =obj.w ;
            try
                doc = invoke(h.Documents, ...           % try to open document as 
                    'Open', od.file_name);           	% existing file
            catch                                       % when this fails 
                doc = invoke(h.Documents,'Add');    % open a new file 
                if numel(regexp(od.file_name,'docx$')) > 0 % test extension
                    file_type = ...                     % docx extension
                        obj.wd2num('wdFormatXMLDocument') ;
                else
                    file_type = ...                 	% doc extension
                        obj.wd2num('wdFormatDocument97');
                end
                SaveAs2(doc,od.file_name,file_type)        % save new file as doc or docx file
            end
        end
        function close_doc(~,doc)                       
            %CLOSE_DOC Close document without saving first
            doc.Saved = true;                           % indicate that it is already saved
            doc.Close() ;                               % close the file
        end
        function saveas_doc(~,doc,varargin)         
            %SAVEAS_DOC Save document (with another name?)
            od = struct(    ...                         % defaults for arguments
                'file_name', doc.FullName  ...      	% file name
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults  
            invoke(doc, 'SaveAs2', od.file_name);       % save it with the derived name
        end
        function save_doc(~,doc)
            %SAVE_DOC Save document (under its old name)
            invoke(doc, 'Save');                        % save document (under its old name)
        end
        function doc = activate_doc(~,doc)
            % ACTIVATE_DOC Make a document the active one
            doc.Activate() ;                            % make a document the active one
        end    
        function doc = get_active_doc(obj)
            % GET_ACTIVE_DOC Retrieve the active document object
            doc = obj.w.ActiveDocument ;                % retrieve the active document
        end   
        function set_saved_doc(~,doc)
            % SET_SAVED_DOC Indicate that the document does not need to be saved
            doc.Saved = 1;                              % indicate that the document does not need to be saved
        end
        function AddParagraph(obj,varargin)
            % ADDPARAGRAPH Add a paragraph at the end of the document
         	od = struct(    ...                         % defaults for arguments
                'color', 'wdAuto'  ...                  % color of paragraph as one of enum WdColorIndex
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;                                 % actxserver handle	 
            par = h.ActiveDocument.Paragraphs.Add() ;   % add a new paragraph at end of document
            par = par.Next() ;                          % point to new paragraph (?? apparently needed)
            par.Range.Style = obj.currentStyle ;        % give the paragraph the current style
            h.Selection.Start = par.Range.Start ;       % set the selection to the whole range
            h.Selection.End = par.Range.End ;           % of the paragraph (next formatting does not work for Range?)
            h.Selection.ClearFormatting() ;             % remove all formatting from the selection
            h.Selection.Style = obj.currentStyle ;      % give the selection the current style
            h.Selection.Font.ColorIndex =  ...          % give the font of the selection the chosen color
                obj.wd2num(od.color) ;
        end %AddParagraph    
        function AddSymbol(obj,varargin)
            % ADDSYMBOL Add a symbol represented by an integer
            % Integer can be found in the Insert/Symbol menu
            % e.g. degree symbol = xB0 = 176           
            od = struct(    ...                         % defaults for arguments
                'symbol',176  ...                       % symbol
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults    
            obj.w.Selection.InsertSymbol(od.symbol);  	% insert in selection           
        end % AddSymbol
        function SelectionSetFont(obj,varargin)
            % SELECTIONSETFONT  Set font attributes for selection
            %  varargin is a structure with one or more of the following fields
            %             defopts     = struct( ...
            %                 'Name',  'Arial'                , ...
            %                 'Size',  9                      , ...
            %                 'Bold',  false                  , ...
            %                 'Italic', false                 , ...
            %                 'Underline', false              , ...
            %                 'StrikeThrough', false          , ...
            %                 'ColorIndex', 'wdAuto'          , ...
            %                 'DoubleStrikeThrough', false 	  , ...
            %                 'Superscript', false            , ...
            %                 'Subscript', false      ) ;
            od = struct() ;
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults                        
            selectionfont = obj.w.Selection.Font ;      % font object of the selection
            fn = fieldnames(od) ;                       % specified options
            for f = fn'                                 % for each of the options
                if strcmpi(f{1},'ColorIndex')           % if this option is related to the ColorIndex
                    cc =od.(f{1}) ;                     % copy the ColorIndex constant
                    selectionfont.ColorIndex = ...      % set the option to the translated constant
                        obj.wd2num(cc) ;
                else
                    selectionfont.(f{1}) = od.(f{1});   % set the option to the specified value
                end
            end
        end % SelectionSetFont
        function SelectionSetAlignment(obj,varargin)
            % SELECTIONSETALIGNMENT  Set alignment for selection
            od = struct(    ...                         % defaults for arguments
                'align','wdAlignParagraphCenter'  ...  	% alignment one of enum WdParagraphAlignment
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults   
            obj.w.Selection.ParagraphFormat. ...        % alignment for selection
                Alignment = obj.wd2num(od.align) ;            
        end % SelectionSetAlignment
        function SelectionInsertText(obj,varargin)
            % SELECTIONINSERTTEXT  Insert text over or besides selection
          	od = struct(    ...                         % defaults for arguments
                'text',  '', ...                        % text to insert
                'lineBreaks', [0,0], ...                % vector with line breaks before and after text
                'color', [], ...                        % color to use as one of enum WdColorIndex
                'fun', 'InsertAfter',  ...            	% function to use (TypeText, InsertBefore or InsertAfter )            
                'collapse',true  ...                    % Collapse section after insert: true (yes: at end ), false (no)
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;                                 % actxserver handle	            
            userOvertype = h.Options.Overtype ;         % save Overtype option
        	h.Options.Overtype = false ;                % Make sure Overtype is turned off.
            switch h.Selection.Type                     % depending on selection type
                case 'wdSelectionIP'                    % insertion point
                    h.Selection.TypeText(od.text)   	% add the text
                    obj.Select({'Left','wdCharacter', ...% and select it
                        numel(od.text),'wdExtend'})
                    insert_done = true ;                % indicate insertion is done
                case 'wdSelectionNormal'                % range selected
                    switch od.fun                     	% insert text with selected function
                        case 'InsertBefore'
                            h.Selection.InsertBefore(od.text);
                        case 'TypeText'
                            h.Selection.Text=od.text;
                        case 'InsertAfter'
                            h.Selection.InsertAfter(od.text);
                    end
                    insert_done = true ;              	% indicate insertion is done
                otherwise
                    insert_done = false ;            	% indicate insertion is not done
            end
            h.Options.Overtype =  userOvertype ;        % reset option to previous value
            if insert_done                              % if insertion was done
                for k = 1:od.lineBreaks(1)
                    h.Selection. ...                    % insert paragraph breaks before insertion
                        InsertParagraphBefore; 
                end
                for k = 1:od.lineBreaks(2)
                    h.Selection. ...                    % insert paragraph breaks after insertion
                        InsertParagraphAfter;
                end
                h.Selection.Style = obj.currentStyle;   % give selection the current style
                if ~isempty(od.color)
                    h.Selection.Font.ColorIndex = ...
                        obj.wd2num(od.color) ;        	% give font the indicated color
                else
                    h.Selection.Font.ColorIndex = ...
                        obj.wd2num('wdAuto') ;          % give font the default color
                end
                if od.collapse
                    h.Selection.Collapse( ...           % collapse selection at end
                        obj.wd2num('wdCollapseEnd'));
                end
            end
        end %SelectionInsertText
        function varargout = InsertTable(obj,varargin)
            % INSERTTABLE Insert a table in the document with caption and bookmark
            od = struct(    ...                         % defaults for arguments
                ... 'dataCell', {'1', '2'; '3', '4'}, ... 	% data
            	'dataCell', [], ...                     % (character) data to insert 
                'lineBreaks', [1,1], ...             	% vector with number of line breaks before and after table
                'bmn', '', ...                          % bookmark name to associate with table
                'captitle', '', ...                     % caption to be given to table
                'cappos', 'wdCaptionPositionBelow', ...	% position of caption: one of enum WdCaptionPosition
                'align', 'wdAlignRowCenter',  ...      	% alignment one of enum WdRowAlignment
                'title', '' ,   ...                     % title of table (in table properties)                   
                'descr', ''     ...                     % description of table (in table properties) 
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;                                 % actxserver handle		            
            [nbRows, nbCols] = size(od.dataCell);            
            for k = 1:od.lineBreaks(1)
                h.Selection.TypeParagraph  ;            % Line breaks before table
            end
            % Create the table and set properties
            newtab = h.ActiveDocument.Tables.Add( ...
                h.Selection.Range, nbRows, nbCols, ...
                obj.wd2num('wdWord9TableBehavior'), ... % Enables AutoFit
                obj.wd2num('wdAutoFitContent'));     	% AutoFitContent
            if numel(od.title) > 0   
                newtab.Title = od.title ;            	% set title when specified
            end
            if numel(od.descr) > 0        
                newtab.Descr = od.descr ;               % set description when specified
            end
            align1 = obj.wd2num(od.align) ;           	% translate table alignment 
            newtab.Rows.Alignment = align1 ;            % apply to tables
            if nargout > 0
                varargout{1} = newtab ;                 % return table object if requested
            end
            obj.SetStyle({obj.normal});                 % set current style to 'Normal'
            % Write data into table
            for r = 1:nbRows
                for c = 1:nbCols
                	obj.SelectionInsertText({od.dataCell{r, c}, ... 	% Write data into current cell
                       	[0, 0],[],'TypeText'});
                    if(r*c == nbRows*nbCols)
                        h.Selection.MoveDown;           % Done, leave the table
                    else 
                        h.Selection.MoveRight;          % Move on to next cell
                    end
                end
            end           
            for k = 1:od.lineBreaks(2)
                h.Selection.TypeParagraph ;             % Line breaks after table
            end
            if numel(od.bmn) > 0                        % if bookmark name is specified
                h.ActiveDocument.Bookmarks.Add ...      % create bookmark for table range
                    (od.bmn,newtab.Range);
            end
            if numel(od.captitle) > 0                  	% if caption is specified
                % Create 'Table' caption below or above table
                r = newtab.Range ;                      % range of this table
                switch od.cappos
                    case 'wdCaptionPositionAbove'
                        dir2cap = 'up' ;
                    case 'wdCaptionPositionBelow'
                        dir2cap = 'down' ;
                    	r.Collapse( ...                 % collapse table range at end
                            obj.wd2num('wdCollapseEnd'));
                end
                r.InsertCaption('Table', ...            % insert caption below or above
                    [' ',od.captitle], obj.wd2num(od.cappos)) 
                r.ParagraphFormat.Alignment = align1 ; 	% caption is aligned with rows
                % Create bookmark for the caption
                if numel(od.bmn) > 0                            % if bookmark name specified
                    newtab.Range.Select;                        % select the table range
                    obj.Select({dir2cap,'wdLine',1,'wdMove'});    % move selection one line up or down
                    obj.Select({'end','wdLine',1,'wdMove'});      % go to end of this line
                    obj.Select({'home','wdLine',1,'wdExtend'});   % extend selection to the start of this line
                    h.ActiveDocument.Bookmarks.Add( ...     % create a bookmark for the selection 
                        [od.bmn,'_caption'], ...
                        h.Selection.Range);
                end
            end
        end % InsertTable
        function varargout = InsertFigure (obj,varargin)
            % INSERTFIGURE Insert a figure in the document with caption and bookmark
               od = struct(    ...                  	% defaults for arguments
                'what', [], ...                         % graphics handle or graphics file name
                'lineBreaks', [1,1], ...             	% vector with number of line breaks before and after figure
                'bmn', '', ...                          % bookmark name to associate with figure
                'captitle', '', ...                     % caption to be given to figure
                'cappos', 'wdCaptionPositionBelow', ...	% position of caption: one of enum WdCaptionPosition
                'align', 'wdAlignParagraphCenter',  ...	% alignment one of enum WdParagraphAlignment
                'perc',  [], ...                        % image width as percentage of document width (0-100)
                'title', '' ,   ...                     % title of figure (in figure properties)                 
                'descr', ''     ...                     % description of figure (in figure properties) 
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;                                 % actxserver handle	
            for k = 1:od.lineBreaks(1)
                h.Selection.TypeParagraph  ;            % Line breaks before figure
            end
            % Insert the figure and set properties
            if ishandle(od.what)                       	% argument is graphics handle
%                 filetype    = 'jpg' ;                   % use jpg format
%                 filename    = sprintf('%s.%s', ...      % generate temp file name
%                     tempname,filetype);
%                 saveas(od.what, filename, filetype) ;   % save image to jpg temp file
               	filetype    = 'emf' ;                   % use jpg format
                filename    = sprintf('%s.%s', ...      % generate temp file name
                    tempname,filetype);
                set(0,'CurrentFigure',od.what)
                print(filename,'-dmeta','-r600')
            else
                filename    = od.what ;               	% use filename as given
            end
            newfig= h.Selection.InlineShapes. ...   % insert the (temporary) graphics file
                AddPicture(filename);
            if ishandle(od.what)
                delete(filename) ;                      % remove temp file
            end
            if numel(od.perc) > 0                    	% when specified                              
                newfig.ScaleHeight = od.perc ;       	% scale height with the given percentage
                newfig.ScaleWidth  = od.perc ;        	% scale width with the given percentage
            end
            if numel(od.title) > 0
                newfig.Title = od.title ;               % set title when specified
            end
            if numel(od.descr) > 0
                newfig.AlternativeText = od.descr ;   	% set description when specified
            end
            if nargout > 0
                varargout{1} = newfig ;
            end
            obj.SetStyle({obj.normal});                   % set current style to 'Normal'
            for k = 1:od.lineBreaks(2)
                h.Selection.TypeParagraph ;             % Line breaks after table
            end
            egraph = h.Selection.Range ;
            egraph.Collapse(obj.wd2num('wdCollapseEnd'));	% point after graph area
            if numel(od.captitle) > 0
                r = newfig.Range ;
                switch od.cappos
                    case 'wdCaptionPositionAbove'
                        r.Collapse( ...                              	% collapse selection at start
                            obj.wd2num('wdCollapseStart'));
                        r.Select
                        obj.Select({'left','wdCharacter',1,'wdExtend'}) 	% select figure
                        h.Selection.InsertCaption('Figure', ...     	% insert caption
                            [' ',od.captitle],obj.wd2num(od.cappos)) ;
                        obj.Select({'home','wdLine',1,'wdExtend'})
                        h.Selection.ParagraphFormat.Alignment = ...     % alignment for caption
                            obj.wd2num(od.align) ;
                    case 'wdCaptionPositionBelow'                        
                        r.Collapse( ...                             	% collapse selection at end
                            obj.wd2num('wdCollapseEnd'));
                        r.Select
                        h.Selection.InsertCaption('Figure', ...     	% insert caption
                            [' ',od.captitle],obj.wd2num(od.cappos)) ;
                        r.Select()
                        h.Selection.TypeParagraph ;
                        h.Selection.ParagraphFormat.Alignment = ...     % alignment for caption
                            obj.wd2num(od.align) ;
                        obj.Select({'end','wdLine',1,'wdMove'})           % move to the end of caption
                        obj.Select({'home','wdLine',1,'wdExtend'});   	% extend selection to start of caption
                end
                if numel(od.bmn) > 0
                    h.ActiveDocument.Bookmarks.Add( ...         % bookmark for caption
                        [od.bmn,'_caption'],h.Selection.Range);                    
                    h.ActiveDocument.Bookmarks.Add( ...         % bookmark for figure
                        od.bmn,newfig.Range);
                end
                newfig.Range.ParagraphFormat.Alignment = ...  	% alignment for figure
                    obj.wd2num(od.align) ;
            else
                if numel(od.bmn) > 0
                    h.ActiveDocument.Bookmarks.Add(od.bmn,newfig.Range);
                end
            end
            egraph.Select() ;                           % set selection to last point of graph area
        end % InsertFigure
        function InsertXRefCaption(obj,varargin)
            % INSERTXREFCAPTION  Insert cross reference to a caption
            od = struct(    ...                         % defaults for arguments
                'bmn', '', ...                          % bookmark name of table, figure, ...
                'reftype', 'wdCaptionTable', ...      	% ReferenceType one of enum WdReferenceKind   
                'refkind', 'wdOnlyLabelAndNumber', ...  % ReferenceKind one of enum WdReferenceKind
                'ashyper', true ...                     % Insert as hyperlink
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;                                 % actxserver handle		
            bkmrk   = [od.bmn,'_caption'] ;         	% bookmark to caption of table
            cap_text = h.ActiveDocument. ...            % caption text
                Bookmarks.Item(bkmrk).Range.Text ;
            reftype = obj.wd2num(od.reftype) ;      	% translated reference type
            all_captions = ...                          % all captions of this reftype
                h.ActiveDocument. ...
                GetCrossReferenceItems(reftype) ;
            [~,ix] = ismember(cap_text,all_captions) ;  % find seq nr of this caption in all captions
            h.Selection.InsertCrossReference( ...
                reftype , ...                           % ReferenceType
                obj.wd2num(od.refkind), ...           	% ReferenceKind
                ix, ...                  				% ReferenceItem = sequence number
                od.ashyper) ;                           % insert as hyperlink ?
        end % InsertXRefCaption
        function SetTableColWidth(obj,varargin)
            % SETTABLECOLWIDTH  Set column widths for table
            od = struct(    ...                         % defaults for arguments
                'columns', [] , ...                     % column number to change ([] is all)
                'widths', [], ...                       % widths in cm or as a percentage
                'width_type', 'wdPreferredWidthPoints'  ...  % width type one of enum WdPreferredWidthType
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;    
            if h.Selection.Information('wdWithInTable') 
                cols    =  h.Selection.Tables. ...      % Columns object for this table
                    Item(1).Columns ;
                nbCols = cols.Count;                    % number of columns
                if numel(od.columns) == 0
                    columns = 1:nbCols;                 % columns to change: all
                else
                    columns = od.columns(:)' ;         	% columns to change: force to row vector 
                end                
                if min(columns) < 1 || max(columns) > nbCols
                     fprintf(['column specification', ...  % warning message
                        ' does not match', ...
                        ' number of columns\n']);
                    return
                end
                if not(ismember(numel(od.widths), ...	% numel(widths) should be 1 or be equal to
                        [1,numel(columns)]))            %   number of specified columns
                    fprintf(['number of elements', ...  % warning message
                        ' in width does not match', ...
                        ' number of columns\n']);
                    return
                end
                if numel(od.widths) == 1
                    widths = repmat(od.widths,1,  ...   % expand a scalar
                        numel(columns))';
                else                    
                    widths = od.widths;                 % copy 
                end
                h.Selection.Tables.Item(1). ...         % disallow AutoFit 
                    AllowAutoFit = false ;
                for i = columns
                    cols.Item(i).PreferredWidthType ...
                        = obj.wd2num(od.width_type) ;
                    switch od.width_type                        
                        case 'wdPreferredWidthAuto'
                            h.Selection.Tables.Item(1). ...     % allow AutoFit
                                AllowAutoFit = true ;
                        case 'wdPreferredWidthPoints'
                            cols.Item(i).PreferredWidth = ...       % convert cm to points
                                widths(i) .* 28.3465;
                        case 'wdPreferredWidthPercent'
                            cols.Item(i).PreferredWidth = ...       % copy percentage
                                widths(i) ;
                    end
                end
            end
        end % SetTableColWidth
        function SetTableColAlign(obj,varargin)
            % SETTABLECOLALIGN  Set alignment for columns for table
            od = struct(    ...                         % defaults for arguments                
                'columns', [] , ...                     % column numbers to change ([] is all)
                'align', []  ...                        % cell array with WdParagraphAlignment constants
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;             
            align1      = obj.wd2num(od.align) ;
            if h.Selection.Information('wdWithInTable')
                nbCols = h.Selection.Tables.Item(1).Columns.Count;
                if numel(od.columns) == 0
                    columns = 1:nbCols;                 % columns to change: all
                else
                    columns = od.columns(:)' ;           % columns to change: force to row vector 
                end                
                if min(columns) < 1 || max(columns) > nbCols
                     fprintf(['column specification', ...  % warning message
                        ' does not match', ...
                        ' number of columns\n']);
                    return
                end
                if not(ismember(numel(align1),[1, ...
                       numel(columns)]))
                    fprintf('number of elements in align does not match number of columns\n')
                    return
                end
                if numel(align1) == 1
                    align1 = repmat(align1,1, ...
                        numel(columns))';
                end
                t = h.Selection.Tables.Item(1) ;
                for i = columns
                    t.Columns.Item(i).Select
                    h.Selection.ParagraphFormat.Alignment = align1(i) ;
                end
                t.Select ;
            end
        end % SetTableColAlign
        function SelectTableColumn(obj,varargin)
            % SELECTTABLECOLUMN  Select a column of a table
            od = struct(    ...                         % defaults for arguments
                'colnr', 1  ...                         % number of the column to select
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;             
            if h.Selection.Information('wdWithInTable')
                nbCols = h.Selection.Tables.Item(1).Columns.Count;
                if (od.colnr < 0) || (od.colnr > nbCols)
                    fprintf('selected column not in table\n')
                    return
                end
                h.Selection.Tables.Item(1).Columns.Item(od.colnr).Select;
            end
        end % SelectTableColumn
        function SelectTableRow(obj,varargin)
            % SELECTTABLEROW  Select a row of a table
            od = struct(    ...                         % defaults for arguments
                'rownr', 1  ...                         % number of the row to select
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;  
            if h.Selection.Information('wdWithInTable')
                nbRows = h.Selection.Tables.Item(1).Rows.Count;
                if (od.rownr < 0) || (od.rownr > nbRows)
                    fprintf('selected row not in table\n')
                    return
                end
                h.Selection.Tables.Item(1).Rows.Item(od.rownr).Select;
            end
        end % SelectTableRow
        function SelectTableCell(obj,varargin)
            % SELECTTABLECELL   Select a cell of a table
            od = struct(    ...                         % defaults for arguments
                'rownr', 1 , ...                      	% row number of the cell to select
                'colnr', 1  ...                         % column number of the cell to select
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;  
            rownr =od.rownr ; colnr = od.colnr ;
            if numel(colnr) == 0
                colnr   = rownr{2} ;
                rownr   = rownr{1} ;
            end
            if h.Selection.Information('wdWithInTable')
                nbRows = h.Selection.Tables.Item(1).Rows.Count;
                nbCols = h.Selection.Tables.Item(1).Columns.Count;
                if (rownr < 0) || (rownr > nbRows) || ...
                        (colnr < 0) || (colnr > nbCols)
                    fprintf('selected cell not in table\n')
                    return
                end
                h.Selection.Tables.Item(1).Cell(rownr,colnr).Select;
            end
        end % SelectTableCell
        function [data,titel, descr] = GetTable(obj,varargin)
            % GETTABLE Get data from current table
            od = struct(    ...                         % defaults for arguments
                'ignoreHeader', true  ...               % ignore row header ?
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;             
            if h.Selection.Information('wdWithInTable') % check cursor is in table
                nbRows = h.Selection.Tables(1).Item(1).Rows.Count - od.ignoreHeader; % ignoreHeader = true (1) or false (0)
                nbCols = h.Selection.Tables(1).Item(1).Columns.Count;
                data = cell(nbRows, nbCols);            % preallocate
                for col = 1:nbCols
                    for row = 1:nbRows
                        cellText = h.Selection.Tables. ...
                            Item(1).Cell(row+od.ignoreHeader,col). ...
                            Range.Text; 
                        data{row,col} = cellText(1:end-2); % end-2 -> ignore line break
                    end
                end
                descr = h.Selection.Tables.Item(1).Descr;
                titel = h.Selection.Tables.Item(1).Title;
            end
        end % GetTable
        function WriteTable(obj,varargin)
            % WRITETABLE Write data to current table
             od = struct(    ...                     	% defaults for arguments
                'data', [] , ...                      	% (character) data to write to table
                'fstRow', 1 , ...                      	% number of the first row to write to
                'fstCol', 1  ...                     	% number of the first column to write to
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;             
            if h.Selection.Information('wdWithInTable') % check cursor is in table
                autofit = h.Selection.Tables.Item(1).AllowAutoFit;
                if autofit  % turn off AutoFit for a faster writing
                    h.Selection.Tables.Item(1).AllowAutoFit = false;
                end
                [dt_ndRows, dt_nbCols] = size(od.data);
                tbl_nbRows = h.Selection.Tables.Item(1).Rows.Count;
                tbl_nbCols = h.Selection.Tables.Item(1).Columns.Count;
                nbCols = min(dt_nbCols, tbl_nbCols-od.fstCol+1);
                ndRows = min(dt_ndRows, tbl_nbRows-od.fstRow+1);
                for col = 1:nbCols
                    for row = 1:ndRows
                        h.Selection.Tables.Item(1). ...
                            Cell(row+od.fstRow-1,col+od.fstCol-1).Range.Text = od.data{row,col};
                    end
                end
             	h.Selection.Tables.Item(1).AllowAutoFit = autofit;
            end
        end % WriteTable
        function SetRowNumb(obj,varargin)
            % SETROWNUMB Change number of rows in current table
            od = struct(    ...                     	% defaults for arguments
                'nbRows', 1   ...                      	% set number of rows to nbRows
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;   
            tbl_nbRows =  h.Selection.Tables.Item(1).Rows.Count;
            row_diff = od.nbRows - tbl_nbRows;
            if ~(row_diff==0)                           % increase or decrease
                if row_diff>0                           % add rows
                    for i=1:row_diff
                        h.Selection.Tables.Item(1).Rows.Add;
                    end
                else                                    % remove rows
                    for i=1:(-row_diff)
                        h.Selection.Tables.Item(1).Rows.Last.Delete;
                    end
                end
            end
        end  % SetRowNumb
        function SetColumnNumb(obj,varargin)
            % SETCOLUMNNUMB Change number of cols in current table
            od = struct(    ...                     	% defaults for arguments
                'nbCols', 1   ...                      	% set number of columns to nbCols
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;   
            tbl_nbCols = h.Selection.Tables.Item(1).Columns.Count;
            col_diff = od.nbCols - tbl_nbCols;
            if ~(col_diff==0)                           % increase or decrease
                if col_diff>0                           % add cols
                    for i=1:col_diff
                        h.Selection.Tables.Item(1).Columns.Add;
                    end
                else          % remove cols
                    for i=1:(-col_diff)
                        h.Selection.Tables.Item(1).Columns.Last.Delete;
                    end
                end                
            end
        end  % SetColumnNumb
        function SetStyle(obj,varargin)
            % SETSTYLE Set current text style, used later by SelectionInsertText
           	od = struct(    ...                     	% defaults for arguments
                'style', obj.normal   ...             	% style to use
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults  
            obj.currentStyle = od.style;
        end % SetStyle
        function SelectionInsertBreak(obj,varargin)
            % SELECTIONINSERTBREAK Insert a break of specified type at start of selection   
             od = struct(    ...                     	% defaults for arguments
                'breaktype',  'wdPageBreak'   ...       % one of enum WdBreakType  
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;   
        	breaktype   = obj.wd2num(od.breaktype) ;  	% translate constant
            h.Selection.HomeKey;                        % goto start of selection
            h.Selection.InsertBreak(breaktype) ;        % insert the break of specified type replacing the selection
            h.Selection.MoveDown ;                      % moves the selection one wdLine down    
        end % SelectionInsertBreak
        function returnV = Goto(obj,varargin)
            % GOTO Jump to specified location in document
          	od = struct(    ...                     	% defaults for arguments
                'what',  'wdGotoBookmark' ,  ...        % one of enum enum WdGoToItem 
                'which', 'wdGotoAbsolute' ,  ...        % one of enum WdGoToDirection  
                'count',  1 ,  ...                      % number of 'what'
                'name',  '' ,  ...                      % name of 'what'
                'delete',  false   ...              	% delete contents indicated selection 
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;   
            which   = obj.wd2num(od.which) ;          	% translate WdGoToDirection
            what    = obj.wd2num(od.what) ;           	% translate WdGoToItem
            if what == -1
                which = [] ; count = [] ;               % ?? otherwise bookmark is not found
            end
            returnV = h.Selection.GoTo(what, which, count, od.name);
            if od.delete
                h.Selection.Delete;
            end
        end % Goto
        function found = FindText(obj, varargin)
            % FINDTEXT Find text in document            
         	od = struct(    ...                     	% defaults for arguments
                'textToFind',  '' ,  ...                % text to find
                'forward', true ,  ...                  % forward (true) or backward (false) 
                'replacement', '' ,  ...                % replacement text
                'find_opts',  struct()   ...           	% find options
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults   
            h = obj.w ;
            defopts     = struct( ...                   % options for Find.Execute
                'wdReplace',  'wdReplaceNone'  , ...  	% replacement options
                'wdFindwrap', 'wdFindContinue' , ...  	% WdFindWrap options
                'MatchCase',  false            , ...
                'MatchWholeWord', false        , ...
                'MatchWildcards', false        , ...
                'MatchSoundsLike', false       , ...
                'MatchAllWordForms', false 	   , ...
                'Format', false ) ;
            opts = wordserver.fn_copy_options( ...      % merge extra options with default ones
                defopts, od.find_opts);
            wdrepcode = obj.wd2num(opts.wdReplace) ;
            wdfwcode = obj.wd2num(opts.wdFindwrap) ;
            h.Selection.Find.ClearFormatting;
            h.Selection.Find.Replacement.ClearFormatting
            found = h.selection.Find.Execute( ...
                od.textToFind, opts.MatchCase, ...
                opts.MatchWholeWord, opts.MatchWildcards, ...
                opts.MatchSoundsLike, opts.MatchAllWordForms, ...
                od.forward, wdfwcode, ...
                opts.Format, od.replacement, wdrepcode) ;
        end % FindText
        function Select(obj,varargin)
            % SELECT Extend or move selection 'right', 'left', 'home' or 'end' , 'up' or 'down' 
            od = struct(    ...                     	% defaults for arguments
                'direction',  'right' ,  ...            % direction ('right', 'left', 'home' or 'end' , 'up', 'down' 
                'unit', 'wdCharacter' ,  ...            % unit one of enum WdUnits
                'count',  1 ,  ...                      % number of 'units'
                'extend', 'wdMove'    ...               % 'wdMove' or 'wdExtend'
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults     
            unit        = obj.wd2num(od.unit) ;
            extend   	= obj.wd2num(od.extend) ;
            switch lower(od.direction)
                case 'left'
                    obj.w.Selection.MoveLeft(unit, od.count, extend);
                case 'right'
                    obj.w.Selection.MoveRight(unit, od.count, extend);
                case 'up'
                    obj.w.Selection.MoveUp(unit, od.count, extend);
                case 'down'
                    obj.w.Selection.MoveDown(unit, od.count, extend);
                case 'home'
                    if ismember(unit,[5 6 9 10])
                        obj.w.Selection.HomeKey(unit, extend);
                    end
                case 'end'
                    if ismember(unit,[5 6 9 10])
                        obj.w.Selection.EndKey(unit, extend);
                    end
            end
        end % Select
        function CreateTOC(obj,varargin)
            % CREATETOC Create the table of contents, list of figures or list of tables
            od = struct(                        ...  	% defaults for arguments
                'type',  'TOC' ,                ...   	% type TOC, Figure or Table
                'text', 'Table of Contents' ,   ...    	% text to print above TOC or list
                'fontstruct', [] ,              ... 	% struct for use in SelectionSetFont to print text
                'tableader', 'wdTabLeaderDots', ...     % enum WdTabLeader
                'include_label', true,          ...   	% include label in list of figure or table
             	'UseHeadingStyles', true,       ...     % forced to false for Figure and Table      
             	'UpperHeadingLevel', 1,         ...     % highest level to include in TOC      
             	'LowerHeadingLevel', 3,         ...     % lowest level to include in TOC       
               	'UseFields', false,             ...              
              	'TableID', '',                  ...                   
               	'RightAlignPageNumbers', true,	...       	 
              	'IncludePageNumbers', true, 	...         
               	'AddedStyles', '',              ...                  
               	'UseHyperlinks', true,          ...            	
               	'HidePageNumbersInWeb', true,	...        
               	'UseOutlineLevels', false       ...  
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults                  
            h = obj.w ;   
            odc = struct2cell(od)' ;                     % convert to cell array
            parms  	= odc(6:end);                       % parameters in all three cases
            if not(strcmpi(od.type,'TOC'))
                parms{1} = false ;                      % no UseHeadingStyles when no TOC
                parms = [{od.type},odc(5),parms(1:end-1)] ; % add parameters for list of figures or tables
            end   
            obj.SelectionInsertText({od.text,[0 2],...      % indicate text to use
                'wdAuto', 'TypeText', false}) ;
            if isstruct(od.fontstruct)
                obj.SelectionSetFont(od.fontstruct);       % apply the structure to the selection
            end
            obj.Select({'end','wdLine',1,'wdMove'});
            obj.Select({'down','wdLine',1,'wdMove'});
            if strcmpi(od.type,'TOC')
                h.ActiveDocument.TablesOfContents.Add( ...
                    h.Selection.Range , ...
                    parms{:} );
                h.ActiveDocument.TablesOfContents.Item(1).TabLeader = ...
                    obj.wd2num(od.tableader) ;      
                h.ActiveDocument.TablesOfContents.Format = ...
                    obj.wd2num('wdIndexIndent') ;
            else
                h.ActiveDocument.TablesOfFigures.Add( ...
                    h.Selection.Range , ...
                    parms{:} );
                h.ActiveDocument.TablesOfFigures.Item(1).TabLeader = ...
                    obj.wd2num(od.tableader) ;     
                h.ActiveDocument.TablesOfFigures.Format = ...
                    obj.wd2num('wdIndexIndent') ;
            end
        end % CreateTOC
        function UpdateTOC(obj,varargin)
            % UPDATETOC Update the Table of Contents
            od = struct(                        ...  	% defaults for arguments
               	'upd_pn_only', false            ...   
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults                  
            h = obj.w ;              
            for i=1:h.ActiveDocument.TablesOfContents.Count
                if od.upd_pn_only
                    h.ActiveDocument.TablesOfContents.Item(i).UpdatePageNumbers
                else
                    h.ActiveDocument.TablesOfContents.Item(i).Update
                end
            end
        end % UpdateTOC
        function AddHeaderFooter(obj,varargin)
            % ADDHEADERFOOTER Add header or footer to the document
            od = struct(    ...                         % defaults for arguments
                'align', 'wdCenter', ...                % alignment one of WdAlignmentTabAlignment
                'infooter', true, ...                   % true when in footer otherwise false
                'pagetxt', 'page ', ...             	% prefix for page number
                'inc_page', true, ...                   % true when page number is to be included
                'inc_tot', true, ...                    % true when total number of pages isto be included
                'inctxt', ' of '  ...               	% prefix for total number
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;                                 % actxserver handle
            if ~strcmp(h.ActiveWindow.View.SplitSpecial, ...
                    'wdPaneNone' )
                h.Panes(2).Close;                       % Make sure the window isn't split
            end           
        	h.ActiveWindow.ActivePane.View.Type  =  ... % Make sure we are in printview
                'wdPrintView';            
            cs = h.Selection ;                          % save selection info
            if od.infooter
               	h.ActiveDocument.Sections.Item(1). ...
                    Footers.Item(obj.wd2num('wdHeaderFooterPrimary')).Range.Select ;
            else
               	h.ActiveDocument.Sections.Item(1). ...
                    Headers.Item(obj.wd2num('wdHeaderFooterPrimary')).Range.Select ;
            end
            s = h.Selection ;
            s.TypeText(od.pagetxt);
            if od.inc_page
                s.Fields.Add( ...,
                    h.Selection.Range, ...
                    -1, ...                          	% wdFieldEmpty
                    'PAGE ', ...
                    true) ;
            end
            if od.inc_tot
                s.TypeText(od.inctxt);
                s.Fields.Add( ...,
                    h.Selection.Range, ...
                    -1, ...                             % wdFieldEmpty
                    'NUMPAGES', ...
                    true) ;
            end
            % Switch back to main document view
            obj.Select({'home','wdLine',1,'wdExtend'}) ;
            obj.SelectionSetAlignment({od.align});  
            %h.ActiveWindow.ActivePane.View.Type  = 'wdPrintView';
            if ~strcmpi(h.ActiveWindow.View.SplitSpecial, ...
                    'wdPaneNone')
                h.ActiveWindow.Panes.Item(2).Close;
            end
            h.ActiveWindow.ActivePane.View.Type  = 'wdPrintView';
            h.ActiveWindow.ActivePane.View.SeekView = 'wdSeekMainDocument' ;    
            cs.Select();
        end % AddHeaderFooter
        function PrintMethods(obj,varargin)
            % PRINTMETHODS Print all available methods for a Word ActiveX object
            %   at the end of the ActiveDocument
            % e.g. hw.PrintMethods('Document.Range','Heading ')
         	od = struct(    ...                      	% defaults for arguments
                'category', 'Application', ...        	% object type for which methods are to be listed
                'headingString', [obj.heading,' 2'] ... % prefix for heading style
                );
            od =wordserver.get_options(od,varargin{:});	% merge arguments with defaults        
            h = obj.w ;                                 % actxserver handle
            obj.SetStyle({od.headingString});
            text = strcat(od.category, '-methods');
            obj.SelectionInsertText({text,[0,1],'',1});
            obj.SetStyle({obj.normal});
            text = sprintf(['Methods to be called from Matlab', ...
                ' as hw.w.%s.MethodName(xxx)'],od.category);
            obj.SelectionInsertText({text,[0,0]});
            text = [' where hw is the variable assigned', ...
                ' to the wordserver instance and the', ...
                ' first parameter "handle" should be omitted:'];
            obj.SelectionInsertText({text, [0,2]});
            category = regexp(od.category,'\.','split') ;
            switch numel(category)
                case 1
                    structMethods = invoke(h.(category{1}));
                case 2
                    structMethods = invoke(h.(category{1}).Item(1).(category{2}));
                otherwise
                    fprintf('Category level too high: nothing reported\n') ;
            end
            cellMethods = struct2cell(structMethods);
            for i = 1:length(cellMethods)
                methodString = cellMethods{i};
                obj.SelectionInsertText({methodString, [0,1],'', 1});
            end
        end % PrintMethods
        function out = wd2num (obj,in)
            % covert Word constant (alpha) to numeric
            [~,ix] = ismember( in, obj.wd_alpha_list) ;
            wd_const_num1 = [NaN,obj.wd_num_list] ;
            out = wd_const_num1(1+ix) ;
        end
    end
    methods (Static, Access = private)
        function h = actxserver2(progID, varargin)
            % Creates a COM Automation server but tries to reuse an existing one
            % Yair Altman  17may2009: Try to reuse an existing COM server instance if possible
            % Han Oostdijk 19dec2016: just packed in function
            try
                h = actxGetRunningServer(progID);
                % no crash so probably succeeded to connect to a running server
            catch
                % Never mind - continue normally to start the COM server and connect to it
                h = actxserver(progID);
            end
        end
        function opt_out = get_options(opt_def,varargin)
            % add (overwrite) option in opt_def by the ones given in varargin          
%{
varargin contains name-value pairs with possible exception
of the first element that then is a structure with options
When the first element is a structure then the name-value pairs
will be added (and possibly overwrite) that structure.
When the first argument is a cell array it is converted to
a structure with no more fields than in opt_def and no more
fields than the number of elements in the cell array.            
If the first element is not a structure then the contents of the
name-value pairs will be copied in an empty structure.

The structure created by varargin will by copied to opt_def
and thereby possibly overwriting elements in opt_def.
The resulting structure is returned as opt_out
%}          
            if numel(varargin) == 0
                opt_out = opt_def ;
            else
                if isstruct(varargin{1})                        	% first argument is a struct
                    opt_in = varargin{1} ;                          % copy struct
                    npargs = varargin(2:end) ;                      % name-value pairs
               	elseif iscell(varargin{1})                          % first argument is a cell array
                    nv     = numel(varargin{1}) ;                   % # element in cell array             
                    f      = fieldnames(opt_def) ;                  % field names of default struct
                    nf     = numel(f) ;                             % # element in default struct
                    nvf    = min([nv,nf]) ;                         % minimum of two lenghts
                    opt_in = cell2struct(varargin{1}(1:nvf), ...    % copy cell to struct using the fieldnames in opt_def
                        f(1:nvf)',2) ;                              % but only the # of elements in common
                    npargs = varargin(2:end) ;                  	% name-value pairs
                else                                                % no struct and no cell array must be name-value pairs
                    opt_in = struct() ;                             % no input argument struct
                    npargs = varargin ;                             % all arguments name-value pairs
                end
                nnp = numel(npargs) ;
                if mod(nnp,2) ~= 0
                    error('get_options: name-value pairs have uneven number of elements')
                end
                i_opt = 1 ;
                while i_opt < nnp
                    try
                        opt_in.(npargs{i_opt}) = npargs{i_opt+1};
                    catch ME
                        error(['getoptions name_value pair could not be added', ...
                            ' to struct: problem in elements %d and %d\n'], i_opt,i_opt+1)
                        rethrow(ME)
                    end
                    i_opt = i_opt + 2 ;
                end
                opt_out = wordserver.fn_copy_options( opt_def, opt_in ) ;
            end
        end
        function opts = fn_copy_options( opts, opt_in )
            % fields in opt_in will overwrite the ones in opts
            fnames = fieldnames(opt_in);
            for fi = 1:numel(fnames)
                fieldname = cell2mat(fnames(fi));
                opts.(fieldname)  = opt_in.(fieldname);
            end
        end
        function defArgs = setOptArgs(a,defArgs)
            empty_a 	= cellfun(@(x)isequal(x,[]),a);	% indicate a that are not specified (empty)
            [defArgs{~empty_a}] = a{~empty_a};        	% replace defaults by non-empty one
        end
        function [wd_alpha,wd_num] = wd_def_const()
            % Word constants translated with obj.wd2num
            wd_const = { ...
                ... enum WdParagraphAlignment
                'wdAlignParagraphCenter' ,      1; ...
                'wdAlignParagraphDistribute' ,  4; ...
                'wdAlignParagraphJustify' ,     3; ...
                'wdAlignParagraphJustifyHi' ,   7; ...
                'wdAlignParagraphJustifyLow' ,  8; ...
                'wdAlignParagraphJustifyMed' ,  5; ...
                'wdAlignParagraphLeft' ,        0; ...
                'wdAlignParagraphRight',        2;  ...
                ... enum WdRowAlignment
                'wdAlignRowCenter',             1;  ...
                'wdAlignRowLeft' ,              0;  ...
                'wdAlignRowRight' ,             2;  ...
                ... enum WdSaveFormat (more in https://msdn.microsoft.com/en-us/library/office/ff839952.aspx)
                'wdFormatDocument97' ,          0;  ...	% Word 97 (doc)
                'wdFormatXMLDocument' ,        12;  ...	% Word 2007 (docx)
                'wdFormatDocumentDefault',     16;  ...	% Word 2007 (docx)
                ... enum WdColorIndex
                'wdAuto' ,                      0;  ...
                'wdBlack' ,                     1;  ...
                'wdBlue' ,                      2;  ...
                'wdBrightGreen' ,               4;  ...
                'wdByAuthor' ,                 -1;  ...
                'wdDarkBlue' ,                  9;  ...
                'wdDarkRed' ,                  13;  ...
                'wdDarkYellow' ,               14;  ...
                'wdGray25' ,                   16;  ...
                'wdGray50' ,                   15;  ...
                'wdGreen' ,                    11;  ...
                'wdNoHighlight' ,           	0;  ...
                'wdPink' ,                      5;  ...
                'wdRed' ,                       6;  ...
                'wdTeal' ,                     10;  ...
                'wdTurquoise' ,                 3;  ...
                'wdViolet' ,                   12;  ...
                'wdWhite' ,                     8;  ...
                'wdYellow',                     7;  ...
                ... enum WdReferenceKind                
                'wdContentText',               -1; ...
                'wdEndnoteNumber',              6; ...
                'wdEndnoteNumberFormatted',    17; ...
                'wdEntireCaption',              2; ...
                'wdFootnoteNumber',             5; ...
                'wdFootnoteNumberFormatted',   16; ...
                'wdNumberFullContext',         -4; ...
                'wdNumberNoContext',           -3; ...
                'wdNumberRelativeContext',     -2; ...
                'wdOnlyCaptionText',            4; ...
                'wdOnlyLabelAndNumber',         3; ...
                'wdPageNumber',                 7; ...
                'wdPosition',                  15; ...
                ... enum WdGoToDirection
                'wdGoToAbsolute',               1;  ...
                'wdGoToFirst',                  1;  ...
                'wdGoToLast',                  -1;  ...
                'wdGoToNext',               	2;  ...
                'wdGoToPrevious',               3;  ...
                'wdGoToRelative',               2;  ...
                ... enum WdGoToItem
                'wdGoToBookmark',              -1;  ...
                'wdGoToComment',                6;  ...
                'wdGoToEndnote',                5;  ...
                'wdGoToEquation',              10;  ...
                'wdGoToField',                  7;  ...
                'wdGoToFootnote',               4;  ...
                'wdGoToGrammaticalError',      14;  ...
                'wdGoToGraphic',                8;  ...
                'wdGoToHeading',               11;  ...
                'wdGoToLine',                   3;  ...
                'wdGoToObject',                 9;  ...
                'wdGoToPage',                   1;  ...
                'wdGoToPercent',                2;  ...
                'wdGoToProofreadingError',     15;  ...
                'wdGoToSection',                0;  ...
                'wdGoToSpellingError',         13;  ...
                'wdGoToTable',                  2;  ...
                ... enum WdReplace
                'wdReplaceNone' ,               0 ;  ...
                'wdReplaceOne' ,                1 ;  ...
                'wdReplaceAll' ,                2 ;  ...
                ... enum WdFindWrap
                'wdFindStop' ,               	0 ;  ...
                'wdFindContinue' ,              1 ;  ...
                'wdFindAsk' ,                   2 ;  ...
                ... enum WdUnits
                'wdCharacter' ,                 1 ;  ...
                'wdWord' ,                      2 ;  ...
                'wdSentence' ,                  3 ;  ...
                'wdParagraph' ,                 4 ;  ...
                'wdLine' ,                      5 ;  ...
                'wdStory' ,                     6 ;  ...
                'wdScreen' ,                    7 ;  ...
                'wdSection' ,                   8 ;  ...
                'wdColumn' ,                    9 ;  ...
                'wdRow' ,                      10 ;  ...
                'wdWindow' ,                   11 ;  ...
                'wdCell' ,                     12 ;  ...
                'wdCharacterFormatting' ,      13 ;  ...
                'wdParagraphFormatting' ,      14 ;  ...
                'wdTable' ,                    15 ;  ...
                'wdItem' ,                     16 ;  ...
                ... enum WdMovementType
                'wdMove' ,                      0 ;  ...
                'wdExtend' ,                    1 ;  ...
                ... enum WdCaptionLabelID
                'wdCaptionFigure' ,            -1 ; ...
                'wdCaptionTable' ,             -2 ; ...
                'wdCaptionEquation',  	       -3 ; ...
                ... enum WdDefaultTableBehavior
                'wdWord8TableBehavior' ,   	    0 ; ... Disables AutoFit
                'wdWord9TableBehavior' ,        1 ; ... Enables AutoFit
                ... enum WdAutoFitBehavior
               	'wdAutoFitContent' ,            1 ; ...
                'wdAutoFitFixed' ,              0 ; ...
                'wdAutoFitWindow',  	        2 ; ...
                ... enum WdBreakType
                'wdColumnBreak',                8; ...
                'wdLineBreak',                  6; ...
                'wdLineBreakClearLeft',         9; ...
                'wdLineBreakClearRight',       10; ...
                'wdPageBreak',                  7; ...
                'wdSectionBreakContinuous',     3; ...
                'wdSectionBreakEvenPage',       4; ...
                'wdSectionBreakNextPage',       2; ...
                'wdSectionBreakOddPage',        5; ...
                'wdTextWrappingBreak',         11; ...
                ... enum WdPreferredWidthType
                'wdPreferredWidthAuto' ,        1; ...
            	'wdPreferredWidthPercent' ,     2; ...
                'wdPreferredWidthPoints' ,      3; ...
                ... enum WdCaptionPosition
                'wdCaptionPositionAbove' ,      0; ...
                'wdCaptionPositionBelow' ,      1; ...
                ... enum WdTabLeader 
                'wdTabLeaderDashes',            2 ; ...
                'wdTabLeaderDots',              1 ; ...
                'wdTabLeaderHeavy',             4 ; ...
                'wdTabLeaderLines',             3 ; ...
                'wdTabLeaderMiddleDot',         5 ; ...
                'wdTabLeaderSpaces',            0 ; ...
                ... enum WdIndexType
                'wdIndexIndent',                0 ; ...
                'wdIndexRunin',                 1 ; ...
                ... enum WdPageNumberAlignment
                'wdAlignPageNumberCenter',      1 ; ...
                'wdAlignPageNumberInside',      3 ; ...
                'wdAlignPageNumberLeft',        0 ; ...
                'wdAlignPageNumberOutside',     4 ; ...
                'wdAlignPageNumberRight',       2 ; ...
                ... enum WdAlignmentTabAlignment 
                'wdCenter',                     1 ; ...
                'wdLeft',                       0 ; ...
                'wdRight',                      2 ; ...
                ... enum WdAlignmentTabRelative
                'wdIndent',                     1 ; ...
                'wdMargin',                     0 ; ...
                ... enum WdHeaderFooterIndex 
                'wdHeaderFooterEvenPages',      3 ; ...
                'wdHeaderFooterFirstPage',      2 ; ...
                'wdHeaderFooterPrimary',        1 ; ...
                ... enum WdSpecialPane 
                'wdPaneComments',              15 ; ...
                'wdPaneCurrentPageFooter',     17 ; ...
                'wdPaneCurrentPageHeader',     16 ; ...
                'wdPaneEndnoteContinuationNotice',      12 ; ...
                'wdPaneEndnoteContinuationSeparator',   13 ; ...
                'wdPaneEndnotes',               8 ; ...
                'wdPaneEndnoteSeparator',      14 ; ...
                'wdPaneEvenPagesFooter',        6 ; ...
                'wdPaneEvenPagesHeader',        3 ; ...
                'wdPaneFirstPageFooter',        5 ; ...
                'wdPaneFirstPageHeader',        2 ; ...
                'wdPaneFootnoteContinuationNotice',      9 ; ...
                'wdPaneFootnoteContinuationSeparator',  10 ; ...
                'wdPaneFootnotes',              7 ; ...
                'wdPaneFootnoteSeparator',     11 ; ...
                'wdPaneNone',               	0 ; ...
                'wdPanePrimaryFooter',          4 ; ...
                'wdPanePrimaryHeader',          1 ; ...
                'wdPaneRevisions',             18 ; ...
                ... enum WdViewType 
                'wdMasterView',                 5 ; ...
                'wdNormalView',                 1 ; ...
                'wdOutlineView',                2 ; ...
                'wdPrintPreview',               4 ; ...
                'wdPrintView',                  3 ; ...
                'wdReadingView',                7 ; ...
                'wdWebView',                    6 ; ...
                 ... enum WdSeekView  
                'wdSeekCurrentPageFooter',      10 ; ...
                'wdSeekCurrentPageHeader',      9 ; ...
                'wdSeekEndnotes',               8 ; ...
                'wdSeekEvenPagesFooter',        6 ; ...
                'wdSeekEvenPagesHeader',        3 ; ...
                'wdSeekFirstPageFooter',        5 ; ...
                'wdSeekFirstPageHeader',        2 ; ...
                'wdSeekFootnotes',              7 ; ...
                'wdSeekMainDocument',           0 ; ...
                'wdSeekPrimaryFooter',          4 ; ...
                'wdSeekPrimaryHeader',          1 ; ...
                ... enum WdCollapseDirection
                'wdCollapseEnd',                0 ; ...
                'wdCollapseStart',              1   ...
                } ;
            wd_alpha    = wd_const(:,1)';
            wd_num      = cell2mat(wd_const(:,2)');
        end
    end
end
