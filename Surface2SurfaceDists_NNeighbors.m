%  Installation:
%
%  - Copy this file into the XTensions folder in the Imaris installation directory
%  - You will find this function in the Image Processing menu
%
%    <CustomTools>
%      <SurpassTab>
%        <SurpassComponent name="bpSurfaces">
%          <Item name="Surface2SurfaceDists_NNeighbors" icon="Matlab" tooltip="Find surface to surface distances and number of nearest neighbors.">
%            <Command>MatlabXT::Surface2SurfaceDists_NNeighbors(%i)</Command>
%          </Item>
%        </SurpassComponent>
%      </SurpassTab>
%    </CustomTools>
% 
%
%  Description:
%   Calculates the distances from surfaces in a Surface Object to all other surfaces within another Surfaces Object.
%   Finds the number of nearest neighbors within a user-defined search
%   radius. Save the results into an Excel Spreadsheet.
%
%  Caveats:
%   At the moment, only works on surface objects, and requires a copy of 
%   Microsoft Excel to run. 
%
%  Author:
%   Adam Fries
% 
%  Date:
%   2018.12.07
%
%


function Surface2SurfaceDists_NNeighbors(aImarisApplicationID)

% turn off silly warning
warning( 'off', 'MATLAB:xlswrite:AddSheet' ) ;

% connect to Imaris interface
if ~isa(aImarisApplicationID, 'Imaris.IApplicationPrxHelper')
    javaaddpath ImarisLib.jar
    vImarisLib = ImarisLib;
    if ischar(aImarisApplicationID)
        aImarisApplicationID = round(str2double(aImarisApplicationID));
    end
    vImarisApplication = vImarisLib.GetApplication(aImarisApplicationID);
else
    vImarisApplication = aImarisApplicationID;
end

% the user has to create a scene with some spots and surface
vSurpassScene = vImarisApplication.GetSurpassScene;
if isequal(vSurpassScene, [])
    msgbox('Please create some Spots and Surface in the Surpass scene!')
    return
end

numT = vImarisApplication.GetDataSet.GetSizeT;


% get the spots and the surface object
vSpots = vImarisApplication.GetFactory.ToSpots(vImarisApplication.GetSurpassSelection);
vSurfaces = vImarisApplication.GetFactory.ToSurfaces(vImarisApplication.GetSurpassSelection);

vSpotsSelected = ~isequal(vSpots, []);
vSurfaceSelected = ~isequal(vSurfaces, []);
if vSpotsSelected
    vParent = vSpots.GetParent;
elseif vSurfaceSelected
    vParent = vSurfaces.GetParent;
else
    vParent = vSurpassScene;
end

% get the spots and surfaces
vSpotsSelection = 1;
vSurfaceSelection = 1;
vNumberOfSpots = 0;
vNumberOfSurfaces = 0;
vSpotsList = [];
vSurfacesList = [];
vSpotsName = {};
vSurfacesName = {};
for vIndex = 1:vParent.GetNumberOfChildren
    vItem = vParent.GetChild(vIndex-1);
    if vImarisApplication.GetFactory.IsSpots(vItem)
        vNumberOfSpots = vNumberOfSpots + 1;
        vSpotsList(vNumberOfSpots) = vIndex;
        vSpotsName{vNumberOfSpots} = char(vItem.GetName);
        
        if vSpotsSelected && isequal(vItem.GetName, vSpots.GetName)
            vSpotsSelection = vNumberOfSpots; 
        end
    elseif vImarisApplication.GetFactory.IsSurfaces(vItem)
        vNumberOfSurfaces = vNumberOfSurfaces + 1;
        vSurfacesList(vNumberOfSurfaces) = vIndex;
        vSurfacesName{vNumberOfSurfaces} = char(vItem.GetName);
        
        if vSurfaceSelected && isequal(vItem.GetName, vSurfaces.GetName)
            vSurfaceSelection = vNumberOfSurfaces;
        end
    end
end

% send error if less than two surfaces are created
if min(vNumberOfSurfaces) == 1
    msgbox('Please create at least two surface objects.')
    return
end

numref = 2;
while numref > 1
    %% should only be 1 reference surface population
    if vNumberOfSurfaces>1
        [refSurf,vOk] = listdlg('ListString',vSurfacesName, ...
            'InitialValue', vSurfaceSelection, 'SelectionMode','multiple', ...
            'ListSize',[300 300], 'Name','Find Surface to Surface distances', ...
            'PromptString',{'Please select the REFERENCE surface:'});
        if vOk<1, return, end
    end
    numref = numel(refSurf);
    if numref > 1
        waitfor(msgbox('Please only choose 1 REFERENCE surface.'))
    end
end

numtarg = 2;
while numtarg > 1
    %% should only be 1 reference surface population
    if vNumberOfSurfaces>1
        [targetSurf,vOk] = listdlg('ListString',vSurfacesName, ...
            'InitialValue', vSurfaceSelection, 'SelectionMode','multiple', ...
            'ListSize',[300 300], 'Name','Find Surface to Surface distances', ...
            'PromptString',{'Please select the TARGET surface:'});
        if vOk<1, return, end
    end
    numtarg = numel(targetSurf);
    if numtarg > 1
        waitfor(msgbox('Please only choose 1 TARGET surface.'))
    end
end



rad = inputdlg({'Please enter the search RADIUS from surface(um):'}, ...
    'Surface 2 Surface Distance',1,{'5'});

[filename, path] = uiputfile('*.xls');

savefile = strcat(path, filename);
%% TODO dialog box for the path search to write the file

vProgressDisplay = waitbar(0,'Finding Surface To Surface Distances');

% compute distance from reference to the targets, should only be 1
% population
vNumberOfSurfacesSelected = numel(targetSurf) + 1;

% don't really need a for loop here, the data just passes through this block once
% for vSurfaceIndex = 1:vNumberOfSurfacesSelected - 1
    % grab the reference surface object
    ritem = vParent.GetChild(vSurfacesList(refSurf) - 1);
    refSurf = vImarisApplication.GetFactory.ToSurfaces(ritem);
    
    % grab the target surface object
    titem = vParent.GetChild(vSurfacesList(targetSurf) - 1);
    targSurf = vImarisApplication.GetFactory.ToSurfaces(titem);
    
    nrefSurf = refSurf.GetNumberOfSurfaces;
    ntargSurf = targSurf.GetNumberOfSurfaces;
    refSurfPosList = zeros(nrefSurf, 3);
    targSurfPosList = zeros(ntargSurf, 3);
    
    
    refSurfID = refSurf.GetIds;
    targSurfID = targSurf.GetIds;
   
    
    % open the excel file for writing
    Excel = actxserver('Excel.Application');
    %File = 'D:\Data\Adam\my_File.xls';  %Make sure you put the whole path
    File = savefile;
    if ~exist(File,'file')
        ExcelWorkbook = Excel.workbooks.Add;
        ExcelWorkbook.SaveAs(File);
        ExcelWorkbook.Close(false);
    end
    ExcelWorkbook = Excel.workbooks.Open(File);

    %% loop through each time point
    targtimearray = zeros(ntargSurf, 1);
    reftimearray = zeros(nrefSurf, 1);
    for n = 0:ntargSurf - 1
        % the time index array for the targets
        targtimearray(n + 1) = targSurf.GetTimeIndex(n);
    end
    for n = 0:nrefSurf - 1
        % the time index arrray for the references
        reftimearray(n + 1) = refSurf.GetTimeIndex(n);
    end
    
    
    %% for every time point: 
    %   get the distances from every reference object
    %   to every other target object. If N is the number of references and 
    %   M is the number of targets, then for each time point, the total
    %   number of distance should be contained in a M x N matrix. The size
    %   of this matrix may change from timepoint to timepoint
    
    
    % initialize the previous array lengths, we haven't started yet, so 
    %   there are defined as 0
    lastrefTlen = 0;
    lasttargTlen = 0;
    currrefTlen = 0;
    currtargTlen = 0;
    for k = 0:numT - 1
        % grab the number of reference and target objects in the current
        % timepoint implied by the length of their time index array
        currrefTlen = currrefTlen + length(reftimearray(reftimearray == k));
        dupelen = length(targtimearray(targtimearray == k));
        currtargTlen = currtargTlen + dupelen;
        
        dupes = ones(dupelen, 1);
        data = [];
        
        %% for every reference object (within a timepoint):
        %   calculate the distance from the reference to every target and 
        %   store the array
        ii = 0;
        for i = lastrefTlen:currrefTlen - 1
            
            refSurfPos = refSurf.GetCenterOfMass(i);
            refSurfPosList(ii+1,:) = refSurf.GetCenterOfMass(i);
            dist = zeros(dupelen, 1);
            
            % distance calculation for every target
            jj = 0;
            tids = zeros(dupelen, 1);
            for j = lasttargTlen:currtargTlen - 1
                targSurfPos = targSurf.GetCenterOfMass(j);
                targSurfPosList(jj+1,:) = targSurf.GetCenterOfMass(j);
                delx = refSurfPos(1) - targSurfPos(1);
                dely = refSurfPos(2) - targSurfPos(2);
                delz = refSurfPos(3) - targSurfPos(3);
                dist(jj+1) = sqrt(delx^2 + dely^2 + delz^2);
                tids(jj+1) = targSurfID(j+1);
                jj = jj + 1;
            end
            
            % assign the reference object IDs, distance average and stddev
            %   and time index
            rids = double(refSurfID(i+1))*dupes;
            distavg = mean(dist)*dupes;
            diststd = std(dist)*dupes;
            times = k*dupes;
             
            
            

            % create the data matrix for each reference object and continue
            %   appending to previous data matrix until all reference
            %   objects are considered. The final data matrix will be the 
            %   excel sheet that gets written
            %%
            nrad = str2double(rad);
            numrad = length(dist(dist<=nrad))*dupes;
            
         
           
            
            datachunk = [times rids tids dist distavg diststd nrad*dupes numrad];
            data = [data ; datachunk];
            ii = ii + 1;
        end

        
        % the current object lengths becomes the previous one as we move
        %   to the next time index
        lastrefTlen = currrefTlen; 
       % currrefTlen = currrefTlen + 
        lasttargTlen = currtargTlen;
            
        % write each matrix of data per time point as a sheet in excel
        data_cells=num2cell(data);  
        col_header={'Time Index', ...
            strcat(char(refSurf.GetName), ' ', ' ID'), ...
            strcat(char(targSurf.GetName), ' ', 'ID'), ...
            strcat('Distances to_', char(targSurf.GetName)), ...
            'Distance Mean', ...
            'Distance Standard Deviation', ...
            'Search Radius', ...
            'Number of Targets within Reference Search Radius'};

        output_matrix=[col_header; data_cells];     
        
        % write the sheet to the excel file
        xlswrite2007(savefile,output_matrix, k+1);     
        
        % update the waitbar according to number of time indices
        waitbar((k+1)/numT);
        
    end 
  
    % save the excel spreadsheet and close out Excel
    ExcelWorkbook.Save
    ExcelWorkbook.Close(false)  % Close Excel workbook.
    Excel.Quit;
    delete(Excel); 


    close(vProgressDisplay);


function colLetter = xlcolumnletter(colNumber)
% Excel formats columns using letters.
% This function returns the letter combination that corresponds to a given
% column number.
% Limited to 702 columns
if( colNumber > 26*27 )
    error('XLCOLUMNLETTER: Requested column number is larger than 702. Need to revise method to work with 3 character columns');
else
    % Start with A-Z letters
    atoz        = char(65:90)';
      % Single character columns are first
      singleChar  = cellstr(atoz);
      % Calculate double character columns
      n           = (1:26)';
      indx        = allcomb(n,n);
      doubleChar  = cellstr(atoz(indx));
      % Concatenate
      xlLetters   = [singleChar;doubleChar];
      % Return requested column
      colLetter   = xlLetters{colNumber};
  end

function A = allcomb(varargin)
% ALLCOMB - All combinations
%    B = ALLCOMB(A1,A2,A3,...,AN) returns all combinations of the elements
%    in the arrays A1, A2, ..., and AN. B is P-by-N matrix where P is the product
%    of the number of elements of the N inputs. 
%    This functionality is also known as the Cartesian Product. The
%    arguments can be numerical and/or characters, or they can be cell arrays.
%
%    Examples:
%       allcomb([1 3 5],[-3 8],[0 1]) % numerical input:
%       % -> [ 1  -3   0
%       %      1  -3   1
%       %      1   8   0
%       %        ...
%       %      5  -3   1
%       %      5   8   1 ] ; % a 12-by-3 array
%
%       allcomb('abc','XY') % character arrays
%       % -> [ aX ; aY ; bX ; bY ; cX ; cY] % a 6-by-2 character array
%
%       allcomb('xy',[65 66]) % a combination -> character output
%       % -> ['xA' ; 'xB' ; 'yA' ; 'yB'] % a 4-by-2 character array
%
%       allcomb({'hello','Bye'},{'Joe', 10:12},{99999 []}) % all cell arrays
%       % -> {  'hello'  'Joe'        [99999]
%       %       'hello'  'Joe'             []
%       %       'hello'  [1x3 double] [99999]
%       %       'hello'  [1x3 double]      []
%       %       'Bye'    'Joe'        [99999]
%       %       'Bye'    'Joe'             []
%       %       'Bye'    [1x3 double] [99999]
%       %       'Bye'    [1x3 double]      [] } ; % a 8-by-3 cell array
%
%    ALLCOMB(..., 'matlab') causes the first column to change fastest which
%    is consistent with matlab indexing. Example: 
%      allcomb(1:2,3:4,5:6,'matlab') 
%      % -> [ 1 3 5 ; 1 4 5 ; 1 3 6 ; ... ; 2 4 6 ]
%
%    If one of the N arguments is empty, ALLCOMB returns a 0-by-N empty array.
%    
%    See also NCHOOSEK, PERMS, NDGRID
%         and NCHOOSE, COMBN, KTHCOMBN (Matlab Central FEX)
% Tested in Matlab R2015a and up
% version 4.2 (apr 2018)
% (c) Jos van der Geest
% email: samelinoa@gmail.com
% History
% 1.1 (feb 2006), removed minor bug when entering empty cell arrays;
%     added option to let the first input run fastest (suggestion by JD)
% 1.2 (jan 2010), using ii as an index on the left-hand for the multiple
%     output by NDGRID. Thanks to Jan Simon, for showing this little trick
% 2.0 (dec 2010). Bruno Luong convinced me that an empty input should
% return an empty output.
% 2.1 (feb 2011). A cell as input argument caused the check on the last
%      argument (specifying the order) to crash.
% 2.2 (jan 2012). removed a superfluous line of code (ischar(..))
% 3.0 (may 2012) removed check for doubles so character arrays are accepted
% 4.0 (feb 2014) added support for cell arrays
% 4.1 (feb 2016) fixed error for cell array input with last argument being
%     'matlab'. Thanks to Richard for pointing this out.
% 4.2 (apr 2018) fixed some grammar mistakes in the help and comments
narginchk(1,Inf) ;
NC = nargin ;
% check if we should flip the order
if ischar(varargin{end}) && (strcmpi(varargin{end}, 'matlab') || strcmpi(varargin{end}, 'john'))
    % based on a suggestion by JD on the FEX
    NC = NC-1 ;
    ii = 1:NC ; % now first argument will change fastest
else
    % default: enter arguments backwards, so last one (AN) is changing fastest
    ii = NC:-1:1 ;
end
args = varargin(1:NC) ;
if any(cellfun('isempty', args)) % check for empty inputs
    warning('ALLCOMB:EmptyInput','One of more empty inputs result in an empty output.') ;
    A = zeros(0, NC) ;
elseif NC == 0 % no inputs
    A = zeros(0,0) ; 
elseif NC == 1 % a single input, nothing to combine
    A = args{1}(:) ; 
else
    isCellInput = cellfun(@iscell, args) ;
    if any(isCellInput)
        if ~all(isCellInput)
            error('ALLCOMB:InvalidCellInput', ...
                'For cell input, all arguments should be cell arrays.') ;
        end
        % for cell input, we use to indices to get all combinations
        ix = cellfun(@(c) 1:numel(c), args, 'un', 0) ;
        
        % flip using ii if last column is changing fastest
        [ix{ii}] = ndgrid(ix{ii}) ;
        
        A = cell(numel(ix{1}), NC) ; % pre-allocate the output
        for k = 1:NC
            % combine
            A(:,k) = reshape(args{k}(ix{k}), [], 1) ;
        end
    else
        % non-cell input, assuming all numerical values or strings
        % flip using ii if last column is changing fastest
        [A{ii}] = ndgrid(args{ii}) ;
        % concatenate
        A = reshape(cat(NC+1,A{:}), [], NC) ;
    end
end

function [success,message]=xlswrite2007(file,data,sheet,range)
%This code increases the speed of the xlswrite and works with Excel 2007
%function when used in loops or multiple times. The problem with the original function
%is that it opens and closes the Excel server every time the function is used.
%To increase the speed I have just edited the original function by removing the 
%server open and close function from the xlswrite function and moved them outside 
%of the function. To use this first run the following code which opens the activex 
%server and checks to see if the file already exists (creates if it doesnt):  
 
% Excel = actxserver('Excel.Application');
% File = 'C:\Folder\File';  %Make sure you put the whole path
% if ~exist(File,'file')
%     ExcelWorkbook = Excel.workbooks.Add;
%     ExcelWorkbook.SaveAs(File)
%     ExcelWorkbook.Close(false);
% end
% ExcelWorkbook = Excel.workbooks.Open(File);
 
%Then run the new xlswrite2007 function as many times as needed
%or in a loop (for example xlswrite2007(File,data,location). 
%Then run the following code to close the activex server:  
 
% ExcelWorkbook.Save
% ExcelWorkbook.Close(false)  % Close Excel workbook.
% Excel.Quit;
% delete(Excel); 
% This is a modern version of xlswrite1 posted by Matt Swartz in 2006, and
% as such most of these comments are copied from his original post
%This works exactly like the original xlswrite function, only many many times faster.

%Excel=evalin('base','Excel');
Excel = evalin('caller', 'Excel');
% Set default values.
Sheet1 = 1;
if nargin < 3
    sheet = Sheet1;
    range = '';
elseif nargin < 4
    range = '';
end
if nargout > 0
    success = true;
    message = struct('message',{''},'identifier',{''});
end
% Handle input.
try
    % handle requested Excel workbook filename.
    if ~isempty(file)
        if ~ischar(file)
            error('MATLAB:xlswrite:InputClass','Filename must be a string.');
        end
        % check for wildcards in filename
        if any(findstr('*', file))
            error('MATLAB:xlswrite:FileName', 'Filename must not contain *.');
        end
        [Directory,file,ext]=fileparts(file);
        if isempty(ext) % add default Excel extension;
            ext = '.xls';
        end
        file = abspath(fullfile(Directory,[file ext]));
        [a1 a2] = fileattrib(file);
        if a1 && ~(a2.UserWrite == 1)
            error('MATLAB:xlswrite:FileReadOnly', 'File cannot be read-only.');
        end
    else % get workbook filename.
        error('MATLAB:xlswrite:EmptyFileName','Filename is empty.');
    end
    % Check for empty input data
    if isempty(data)
        error('MATLAB:xlswrite:EmptyInput','Input array is empty.');
    end
    % Check for N-D array input data
    if ndims(data)>2
        error('MATLAB:xlswrite:InputDimension',...
            'Dimension of input array cannot be higher than two.');
    end
    % Check class of input data
    if ~(iscell(data) || isnumeric(data) || ischar(data)) && ~islogical(data)
        error('MATLAB:xlswrite:InputClass',...
            'Input data must be a numeric, cell, or logical array.');
    end
    % convert input to cell array of data.
     if iscell(data)
        A=data;
     else
         A=num2cell(data);
     end
    if nargin > 2
        % Verify class of sheet parameter.
        if ~(ischar(sheet) || (isnumeric(sheet) && sheet > 0))
            error('MATLAB:xlswrite:InputClass',...
                'Sheet argument must be a string or a whole number greater than 0.');
        end
        if isempty(sheet)
            sheet = Sheet1;
        end
        % parse REGION into sheet and range.
        % Parse sheet and range strings.
        if ischar(sheet) && ~isempty(strfind(sheet,':'))
            range = sheet; % only range was specified.
            sheet = Sheet1;% Use default sheet.
        elseif ~ischar(range)
            error('MATLAB:xlswrite:InputClass',...
                'Range argument must be a string in Excel A1 notation.');
        end
    end
catch exception
    if ~isempty(nargchk(2,4,nargin))
        error('MATLAB:xlswrite:InputArguments',nargchk(2,4,nargin));
    else
        success = false;
        message = exceptionHandler(nargout, exception);
    end
    return;
end
%------------------------------------------------------------------------------
% Attempt to start Excel as ActiveX server.
try
    %Excel = actxserver('Excel.Application');
catch exception %#ok<NASGU>
    warning('MATLAB:xlswrite:NoCOMServer',...
        ['Could not start Excel server for export.\n' ...
        'XLSWRITE will attempt to write file in CSV format.']);
    if nargout > 0
        [message.message,message.identifier] = lastwarn;
    end
    % write data as CSV file, that is, comma delimited.
    file = regexprep(file,'(\.xls)$','.csv'); 
    try
        dlmwrite(file,data,','); % write data.
    catch exception
        exceptionNew = MException('MATLAB:xlswrite:dlmwrite', 'An error occurred on data export in CSV format.');
        exceptionNew = exceptionNew.addCause(exception);
        if nargout == 0
            % Throw error.
            throw(exceptionNew);
        else
            success = false;
            message.message = exceptionNew.getReport;
            message.identifier = exceptionNew.identifier;
        end
    end
    return;
end
%------------------------------------------------------------------------------
try
    % Construct range string
    if isempty(strfind(range,':'))
        % Range was partly specified or not at all. Calculate range.
        [m,n] = size(A);
        range = calcrange(range,m,n);
    end
catch exception
    success = false;
    message = exceptionHandler(nargout, exception);
    return;
end
%------------------------------------------------------------------------------
try
    bCreated = false;
    if ~exist(file,'file')
        % Create new workbook.  
        bCreated = true;
        %This is in place because in the presence of a Google Desktop
        %Search installation, calling Add, and then SaveAs after adding data,
        %to create a new Excel file, will leave an Excel process hanging.  
        %This workaround prevents it from happening, by creating a blank file,
        %and saving it.  It can then be opened with Open.
         ExcelWorkbook = Excel.workbooks.Add;
         ExcelWorkbook.SaveAs(file)
         ExcelWorkbook.Close(false);
    end
    
    %Open file
%      ExcelWorkbook = Excel.workbooks.Open(file);
%     if ExcelWorkbook.ReadOnly ~= 0
%         %This means the file is probably open in another process.
%         error('MATLAB:xlswrite:LockedFile', 'The file %s is not writable.  It may be locked by another process.', file);
%     end
    try % select region.
        % Activate indicated worksheet.
        message = activate_sheet(Excel,sheet);
        % Select range in worksheet.
        Select(Range(Excel,sprintf('%s',range)));
    catch exceptionInner % Throw data range error.
        throw(MException('MATLAB:xlswrite:SelectDataRange', sprintf('Excel returned: %s.', exceptionInner.message))); 
    end
    % Export data to selected region.
    set(Excel.selection,'Value',A);
%     ExcelWorkbook.Save
%     ExcelWorkbook.Close(false)  % Close Excel workbook.
%     Excel.Quit;
catch exception
    try
%         ExcelWorkbook.Close(false);    % Close Excel workbook.
    catch
    end
%     Excel.Quit;
%     delete(Excel);                 % Terminate Excel server.
    if (bCreated && exist(file, 'file') == 2)
        delete(file);
    end
    success = false;
%     message = exceptionHandler(nargout, exception);
end
%--------------------------------------------------------------------------
function message = activate_sheet(Excel,Sheet)
% Activate specified worksheet in workbook.
% Initialise worksheet object
WorkSheets = Excel.sheets;
message = struct('message',{''},'identifier',{''});
% Get name of specified worksheet from workbook
try
    TargetSheet = get(WorkSheets,'item',Sheet);
catch exception  %#ok<NASGU>
    % Worksheet does not exist. Add worksheet.
    TargetSheet = addsheet(WorkSheets,Sheet);
    warning('MATLAB:xlswrite:AddSheet','Added specified worksheet.');
    if nargout > 0
        [message.message,message.identifier] = lastwarn;
    end
end
% activate worksheet
Activate(TargetSheet);
%------------------------------------------------------------------------------
function newsheet = addsheet(WorkSheets,Sheet)
% Add new worksheet, Sheet into worsheet collection, WorkSheets.
if isnumeric(Sheet)
    % iteratively add worksheet by index until number of sheets == Sheet.
    while WorkSheets.Count < Sheet
        % find last sheet in worksheet collection
        lastsheet = WorkSheets.Item(WorkSheets.Count);
        newsheet = WorkSheets.Add([],lastsheet);
    end
else
    % add worksheet by name.
    % find last sheet in worksheet collection
    lastsheet = WorkSheets.Item(WorkSheets.Count);
    newsheet = WorkSheets.Add([],lastsheet);
end
% If Sheet is a string, rename new sheet to this string.
if ischar(Sheet)
    set(newsheet,'Name',Sheet);
end
%------------------------------------------------------------------------------
function [absolutepath]=abspath(partialpath)
% parse partial path into path parts
[pathname filename ext] = fileparts(partialpath);
% no path qualification is present in partial path; assume parent is pwd, except
% when path string starts with '~' or is identical to '~'.
if isempty(pathname) && isempty(strmatch('~',partialpath))
    Directory = pwd;
elseif isempty(regexp(partialpath,'(.:|\\\\)','once')) && ...
        isempty(strmatch('/',partialpath)) && ...
        isempty(strmatch('~',partialpath));
    % path did not start with any of drive name, UNC path or '~'.
    Directory = [pwd,filesep,pathname];
else
    % path content present in partial path; assume relative to current directory,
    % or absolute.
    Directory = pathname;
end
% construct absulute filename
absolutepath = fullfile(Directory,[filename,ext]);
%------------------------------------------------------------------------------
function range = calcrange(range,m,n)
% Calculate full target range, in Excel A1 notation, to include array of size
% m x n
range = upper(range);
cols = isletter(range);
rows = ~cols;
% Construct first row.
if ~any(rows)
    firstrow = 1; % Default row.
else
    firstrow = str2double(range(rows)); % from range input.
end
% Construct first column.
if ~any(cols)
    firstcol = 'A'; % Default column.
else
    firstcol = range(cols); % from range input.
end
try
    lastrow = num2str(firstrow+m-1);   % Construct last row as a string.
    firstrow = num2str(firstrow);      % Convert first row to string image.
    lastcol = dec2base27(base27dec(firstcol)+n-1); % Construct last column.
    range = [firstcol firstrow ':' lastcol lastrow]; % Final range string.
catch exception  %#ok<NASGU>
    error('MATLAB:xlswrite:CalculateRange', 'Invalid data range: %s.', range);
end
%----------------------------------------------------------------------
function string = index_to_string(index, first_in_range, digits)
letters = 'A':'Z';
working_index = index - first_in_range;
outputs = cell(1,digits);
[outputs{1:digits}] = ind2sub(repmat(26,1,digits), working_index);
string = fliplr(letters([outputs{:}]));
%----------------------------------------------------------------------
function [digits first_in_range] = calculate_range(num_to_convert)
digits = 1;
first_in_range = 0;
current_sum = 26;
while num_to_convert > current_sum
    digits = digits + 1;
    first_in_range = current_sum;
    current_sum = first_in_range + 26.^digits;
end
%------------------------------------------------------------------------------
function s = dec2base27(d)
%   DEC2BASE27(D) returns the representation of D as a string in base 27,
%   expressed as 'A'..'Z', 'AA','AB'...'AZ', and so on. Note, there is no zero
%   digit, so strictly we have hybrid base26, base27 number system.  D must be a
%   negative integer bigger than 0 and smaller than 2^52.
%
%   Examples
%       dec2base(1) returns 'A'
%       dec2base(26) returns 'Z'
%       dec2base(27) returns 'AA'
%-----------------------------------------------------------------------------
d = d(:);
if d ~= floor(d) || any(d(:) < 0) || any(d(:) > 1/eps)
    error('MATLAB:xlswrite:Dec2BaseInput',...
        'D must be an integer, 0 <= D <= 2^52.');
end
[num_digits begin] = calculate_range(d);
s = index_to_string(d, begin, num_digits);
%------------------------------------------------------------------------------
function d = base27dec(s)
%   BASE27DEC(S) returns the decimal of string S which represents a number in
%   base 27, expressed as 'A'..'Z', 'AA','AB'...'AZ', and so on. Note, there is
%   no zero so strictly we have hybrid base26, base27 number system.
%
%   Examples
%       base27dec('A') returns 1
%       base27dec('Z') returns 26
%       base27dec('IV') returns 256
%-----------------------------------------------------------------------------
if length(s) == 1
   d = s(1) -'A' + 1;
else
    cumulative = 0;
    for i = 1:numel(s)-1
        cumulative = cumulative + 26.^i;
    end
    indexes_fliped = 1 + s - 'A';
    indexes = fliplr(indexes_fliped);
    indexes_in_cells = mat2cell(indexes, 1, ones(1,numel(indexes)));
    d = cumulative + sub2ind(repmat(26, 1,numel(s)), indexes_in_cells{:});
end
%-------------------------------------------------------------------------------
function messageStruct = exceptionHandler(nArgs, exception)
    if nArgs == 0
        throwAsCaller(exception);  	   
    else
        messageStruct.message = exception.message;       
        messageStruct.identifier = exception.identifier;
    end
