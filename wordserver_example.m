%%

%% wordserver_example 
% This example shows some of the features of the wordserver class.
% First a table and a figure are created that are exported to an empty Microsoft Word document.
% Then some paragraphs with text and references to the figure and table are added to the document.
% Also is shown how sections in the document can be found and be given a different color or font.

%% create a table to write to document
x       = (0:pi/12:pi)' ;                               % create some x-values for table
f1      = @(x) sin(x);                                  % function handle for sine function
f2      = @(x) cos(x);                                  % function handle for cosine function
table_n = [x,f1(x),f2(x)] ;                             % create table with function values (numeric)
table_f = arrayfun (@(x) sprintf('%9.6f',x), ...        % same table but now formatted 
    table_n,'Uniform',false) ;
table_f = [{'x','sin(x)','cos(x)'}; table_f] ;          % with headerline ready for copying to word document
%% create a figure to write to document
close all                                               % close all existing figures
h = figure() ;                                          % create new empty figure 
[X,Y] = meshgrid(1:0.5:10,1:20);
Z = sin(X) + cos(Y);
surf(X,Y,Z)
%% start activex server
int_options = struct( ...                               % default constants
                'normal', 'Normal') ;                   % for default style
hw      = wordserver(int_options) ;                   	% start the word activex server 
%% set visibility in three ways
hw.set_visible('visible',true);                     	% show/hide the activity in word (name-value pair method)
hw.set_visible({false});                                % show/hide the activity in word (cell array method)
hw.set_visible(struct('visible',true));             	% show/hide the activity in word (struct methode)
%% create a new docx file
newfile1 = [tempname,'.docx'] ;                        	% (temporary) name of file in the temp folder
doc1    = hw.open_doc({newfile1}) ;                    	% open a document
% hw.activate_doc(doc1) ;                               % with one open document for server doc1 is automatically the active document
%% give the document a footer (with page indication) and header
hw.AddHeaderFooter() ;                                  % all defaults: centered footer with 'page x of xx'
header_info = struct('pagetxt', ...                     % centered header with 'example for wordserver'
    'example for wordserver', 'infooter', false, ...    % not in footer (so header)
    'inc_page', false, 'inc_tot', false);               % do not include page info
hw.AddHeaderFooter(header_info)
%% insert the figure we created
hw.AddParagraph() ;                                     % add a new paragraph at the end of the document
newfig  = hw.InsertFigure( ...                          % insert the figure that was created 
    {h, ...                                          	% what: graphics handle or graphics file name (with full path)
    [0,2] , ...                                      	% line breaks before and after
    'bm_fig1', ...                                   	% bookmark name
    'my caption voor fig1', ...                        	% caption title
    'wdCaptionPositionBelow', ...                       % caption position 'wdCaptionPositionBelow',
    'wdAlignParagraphCenter', ...                       % horizontal figure alignment
    65}   ...                                           % image width as percentage of document width (0-100)
    );
%% insert a table with the contents that was provided
% hw.AddParagraph();                                 	% add a new paragraph at the end of the document
hw.SelectionInsertBreak() ;                             % force a page break
newtab = hw.InsertTable( ...                               % add a table to the paragraph with
    {table_f, ...                                     	% table contents
    [1 1], ...                                          % line breaks before and after
    'bm_tab1', ...                                      % bookmark name
    'function values for sin(x) and cos(x)', ...        % caption for table
    'wdCaptionPositionBelow', ...                       % position of caption  (alternative 'wdCaptionPositionAbove')
    'wdAlignRowCenter'} ...                          	% horizontal table alignment (alternative 'wdAlignRowLeft' 'wdAlignRowRight')
    );
%% set widths and alignment of table columns
hw.Goto({'wdGoToBookmark','wdGoToFirst',[],'bm_tab1'});	% use bookmark to select table
hw.SetTableColWidth({[] ,[5,4,3]});                       % set widths of columns (in cm) of all columns 
hw.SetTableColAlign('align',{'wdAlignParagraphLeft',...	% set alignment of all columns
    'wdAlignParagraphCenter','wdAlignParagraphRight'});
%% reformat header line of table
hw.SelectTableRow({1});                               	% select first row of table (header line)
hw.SelectionSetFont(struct('Name','Arial', ...          % specify font characteristics of this line
    'Size',8,'ColorIndex','wdRed', ...
    'Bold',true,'Italic',true))
hw.SelectionSetAlignment({'wdAlignParagraphCenter'}) ; 	% center the headers in their cells

%% go to start of document and insert text with references
hw.Select({'home','wdStory',[],'wdMove'});             	% set selection to start of document
hw.SetStyle({'Heading 1'}) ;                           	% indicate style to use
hw.SelectionInsertText({'Example of use for '}) ;     	% insert text with indicated style
%                                                     	%   using all defaults (see next section for expanded example)
hw.SelectionInsertText({'wordserver ',[0 1],'wdRed'}); 	% insert text with indicated style , paragraph after
%                                                       %   and specified color
hw.SetStyle() ;                                         % reset style to standard: 'normal'
mytext = ['In this document we show how to insert ',... % define text to write
    'in a word document tables and figures that are produced by MATLAB. ', ...
    'We include a reference to the first plot (' ];
hw.SelectionInsertText({mytext}) ;                    	% insert text with 'normal' style and all defaults 
hw.InsertXRefCaption({'bm_fig1','wdCaptionFigure'}) ;  	% insert reference to the figure with bookmark 'bm_fig1'
%                                                       %       (insert only label and number)
hw.SelectionInsertText({' on page '});                	% insert text with 'normal' style and all defaults 
hw.InsertXRefCaption({'bm_fig1','wdCaptionFigure', ...  % insert reference to the figure with bookmark 'bm_fig1' 
    'wdPageNumber'}) ;                                	%       (insert only page number)
hw.SelectionInsertText({') and the first table ('});   	% insert text with 'normal' style and all defaults 
hw.InsertXRefCaption({'bm_tab1','wdCaptionTable'}) ;  	% insert reference to the table with bookmark 'bm_tab1'
%                                                     	% (insert only label and number)
hw.SelectionInsertText({' on page '});                	% insert text with 'normal' style and all defaults 
hw.InsertXRefCaption({'bm_tab1','wdCaptionTable', ... 	% insert reference to the table with bookmark 'bm_tab1'
    'wdPageNumber'}) ;                                	%       (insert only page number)
hw.SelectionInsertText({').',[0 1]});                	% insert text with 'normal' style and paragraph end
%% insert paragraph 
hw.AddParagraph();                                      % add a new paragraph at the end of the document
hw.SelectionInsertText( ...                             % insert text with 'normal' style and (specified) defaults
    {'We call an angle of 90', ...                              % text
    [0 0], ...                                          % linebreaks before and after
    'wdAuto', ...                                       % color
    'TypeText', ...                                     % TypeText, InsertBefore, InsertAfter
    true});                                            	% Collapse section after insert: true (yes: at end ), false (no)
hw.AddSymbol('symbol',176) ;                                     % insert the 'degree' symbol
hw.SelectionInsertText({' a right angle.', ...        	% insert text with 'normal' style, no paragraph breaks
    [0 1], 'wdGreen'});                                	% and text in green
found = hw.FindText({'90',false});                    	% find the previous '90'
hw.Select({'right','wdCharacter',1,'wdExtend'});      	% extend the selection with 1 character to the right (to include the degree symbol)
format90 = struct('Size', 12, 'Bold', true, ...         % some font characteristics
    'ColorIndex', 'wdRed') ; 
hw.SelectionSetFont(format90);                          % and use these for the selection.
%% insert overview of methods 
hw.Select({'end','wdStory',[],'wdMove'});            	% set selection to end of document
hw.SelectionInsertBreak() ;                             % force a page break
hw.PrintMethods('category', 'Application')              % print the methods for Document.Range
%% insert Table of Contents
hw.Select({'home','wdStory',[],'wdMove'});            	% set selection to start of document
hw.CreateTOC({'Table','List of Tables', ...             % insert Table of Tables
     	struct('Bold',true)}) ;
hw.Select({'home','wdStory',[],'wdMove'});             	% set selection to start of document
hw.CreateTOC({'Figure','List of Figures', ...           % insert Table of Figures
      	struct('Bold',true)}) ; 
hw.Select({'home','wdStory',[],'wdMove'});             	% set selection to start of document
hw.CreateTOC('fontstruct', ...                          % insert Table of Contents
        struct('Bold',true,'ColorIndex', 'wdRed')) ;
%% save and close document
% hw.save_doc(doc1);                                	% save the document with temp name
hw.saveas_doc(doc1,  'file_name' , ...                  % also save with another name
     fullfile(cd(), 'example.docx') );
hw.close_doc(doc1);                                     % close the document
%% close actxserver
hw.quit()                                               % close the actxserver for Microsoft Word
clear('hw')                                             % free all memory for the actxserver
%% delete work document
dos(sprintf('del "%s" ',newfile1));