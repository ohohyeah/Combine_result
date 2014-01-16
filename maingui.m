function varargout = maingui(varargin)
% MAINGUI M-file for maingui.fig
%      MAINGUI, by itself, creates a new MAINGUI or raises the existing
%      singleton*.
%
%      H = MAINGUI returns the handle to a new MAINGUI or the handle to
%      the existing singleton*.
%
%      MAINGUI('CALLBACK',hObject,eventData,handles,...) calls the local
%      function named CALLBACK in MAINGUI.M with the given input arguments.
%
%      MAINGUI('Property','Value',...) creates a new MAINGUI or raises the
%      existing singleton*.  Starting from the left, property value pairs are
%      applied to the GUI before maingui_OpeningFcn gets called.  An
%      unrecognized property name or invalid value makes property application
%      stop.  All inputs are passed to maingui_OpeningFcn via varargin.
%
%      *See GUI Options on GUIDE's Tools menu.  Choose "GUI allows only one
%      instance to run (singleton)".
%
% See also: GUIDE, GUIDATA, GUIHANDLES

% Edit the above text to modify the response to help maingui

% Last Modified by GUIDE v2.5 23-Dec-2013 09:00:08

% Begin initialization code - DO NOT EDIT
gui_Singleton = 1;
gui_State = struct('gui_Name',       mfilename, ...
                   'gui_Singleton',  gui_Singleton, ...
                   'gui_OpeningFcn', @maingui_OpeningFcn, ...
                   'gui_OutputFcn',  @maingui_OutputFcn, ...
                   'gui_LayoutFcn',  [] , ...
                   'gui_Callback',   []);
if nargin && ischar(varargin{1})
    gui_State.gui_Callback = str2func(varargin{1});
end

if nargout
    [varargout{1:nargout}] = gui_mainfcn(gui_State, varargin{:});
else
    gui_mainfcn(gui_State, varargin{:});
    
   % init % 初始化

end
% End initialization code - DO NOT EDIT


% --- Executes just before maingui is made visible.
function maingui_OpeningFcn(hObject, eventdata, handles, varargin)
% This function has no output args, see OutputFcn.
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)
% varargin   command line arguments to maingui (see VARARGIN)

% Choose default command line output for maingui
handles.output = hObject;

% Update handles structure
guidata(hObject, handles);

% UIWAIT makes maingui wait for user response (see UIRESUME)
% uiwait(handles.figure1);


% --- Outputs from this function are returned to the command line.
function varargout = maingui_OutputFcn(hObject, eventdata, handles) 
% varargout  cell array for returning output args (see VARARGOUT);
% hObject    handle to figure
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Get default command line output from handles structure
varargout{1} = handles.output;



% --- Executes on button press in pushbutton1.
function pushbutton1_Callback(hObject, eventdata, handles)
clc;
folder_name = uigetdir;

outputname = strcat(folder_name, '\combined_result.xlsx');
%rename the sheets
sheetnames = {'record','redox_mean','redox_sigma','redox_SE'};
xlsheets(sheetnames,outputname);

attribute = {'date','avg_gray','avg_red','avg_green','avg_blue','std_gray','std_red','std_green','std_blue','s_std_gray','s_std_red','s_std_green','s_std_blue','SEM_gray','SEM_red','SEM_green','SEM_blue','SD_gray','SD_red','SD_green','SD_blue','number of pixels'};
attribute_others ={'date','L_ori_redox','R_ori_redox', 'L_nor_redox', 'R_nor_redox', 'R/L_ori', 'R/L_nor'};

%if(exist(outputname,'file') == 0)
    xlswrite(outputname, attribute,'A1:V1',1);
    xlswrite(outputname, attribute_others,2,'A1:G1');
    xlswrite(outputname, attribute_others,3,'A1:G1');
    xlswrite(outputname, attribute_others,4,'A1:G1');
    disp('start writing ');
%end
%folder是整理過的動物資料編號
datefiles = dir(folder_name);

%mean sigma SE 之column編號
column_constant = [2 10 14];
%L R_normalizedL   L R_ori 在excel檔案�堶捲臚@次出現的時候
row_constant = [5 9 6 10];

%每次與上一次差9
diff_constant = [9 9 9 9];

for i = 3 : length(datefiles)-1
    currentdate = strcat(folder_name , '\', datefiles(i,1).name);
    %當日底下的資料
    disp(datefiles(i,1).name);
    if(datefiles(i,1).isdir)
        %put the date name of current file in the attribute 'record'
        xlsappend(outputname, {datefiles(i,1).name}, 1);
    end
    imagefiles = dir(currentdate);
    
    for j = 3:length(imagefiles)
        if (imagefiles(j,1).isdir)
            %read the result_new.xls of ezch image
           filename = strcat(currentdate,'\',imagefiles(j,1).name, '\','result_new.xls'); 
           if(exist(filename,'file') ==2 )
               %read the 'record' sheet
               [num,txt,raw] = xlsread(filename,1);
               xlsappend(outputname, raw, 1 );
               disp(imagefiles(j,1).name);
           end
            
        end
    end
    
    %after all data in the same date have done
     [num,txt,raw] = xlsread(outputname,1);
     temp_num = [raw(row_constant(3),column_constant(1)), raw(row_constant(4),column_constant(1)) raw(row_constant(1),column_constant(1)), raw(row_constant(2),column_constant(1)) ];
     ratio_ori =(num(row_constant(4)-1,column_constant(1)) /num(row_constant(3)-1,column_constant(1)));
     ratio_nor =(num(row_constant(2)-1,column_constant(1)) /num(row_constant(1)-1,column_constant(1)));
     temp_write = [{datefiles(i,1).name},temp_num,ratio_ori, ratio_nor];
     xlsappend(outputname, temp_write, 2 );
           
     temp_num = [raw(row_constant(3),column_constant(2)), raw(row_constant(4),column_constant(2)) raw(row_constant(1),column_constant(2)), raw(row_constant(2),column_constant(2)) ];
     ratio_ori =(num(row_constant(4)-1,column_constant(2)) /num(row_constant(3)-1,column_constant(2)));
     ratio_nor =(num(row_constant(2)-1,column_constant(2)) /num(row_constant(1)-1,column_constant(2)));
     temp_write = [{datefiles(i,1).name},temp_num,ratio_ori, ratio_nor];
     xlsappend(outputname, temp_write, 3 );
     
     temp_num = [raw(row_constant(3),column_constant(3)), raw(row_constant(4),column_constant(3)) raw(row_constant(1),column_constant(3)), raw(row_constant(2),column_constant(3)) ];
      ratio_ori =(num(row_constant(4)-1,column_constant(3)) /num(row_constant(3)-1,column_constant(3)));
     ratio_nor =(num(row_constant(2)-1,column_constant(3)) /num(row_constant(1)-1,column_constant(3)));
     temp_write = [{datefiles(i,1).name},temp_num,ratio_ori, ratio_nor];
     xlsappend(outputname, temp_write, 4 );
     
     row_constant =row_constant + diff_constant;
    
     
end
disp('finished');




function edit1_Callback(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    structure with handles and user data (see GUIDATA)

% Hints: get(hObject,'String') returns contents of edit1 as text
%        str2double(get(hObject,'String')) returns contents of edit1 as a double


% --- Executes during object creation, after setting all properties.
function edit1_CreateFcn(hObject, eventdata, handles)
% hObject    handle to edit1 (see GCBO)
% eventdata  reserved - to be defined in a future version of MATLAB
% handles    empty - handles not created until after all CreateFcns called

% Hint: edit controls usually have a white background on Windows.
%       See ISPC and COMPUTER.
if ispc && isequal(get(hObject,'BackgroundColor'), get(0,'defaultUicontrolBackgroundColor'))
    set(hObject,'BackgroundColor','white');
end
