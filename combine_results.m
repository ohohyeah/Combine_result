function [ ] = combine_results( folder_name )

outputname = strcat(folder_name, '\combined_result_v2.xlsx');
%rename the sheets
sheetnames = {'record','redox_mean','redox_sigma','redox_SE', 'intensity_mean', 'intensity_sigma', 'intensity_SE'};
xlsheets(sheetnames,outputname);

attribute = {'date','avg_gray','avg_red','avg_green','avg_blue',...
    'std_gray','std_red','std_green','std_blue',...
    's_std_gray','s_std_red','s_std_green','s_std_blue',...
    'SEM_gray','SEM_red','SEM_green','SEM_blue',...
    'SD_gray','SD_red','SD_green','SD_blue','number of pixels'};
attribute_redox ={'date','L_ori_redox','R_ori_redox', 'L_nor_redox', 'R_nor_redox', 'R/L_ori', 'R/L_nor'};
attribute_raw  = {'date', 'L_335', 'R_335', 'L_460', 'R_460'};
%if(exist(outputname,'file') == 0)

xlswrite(outputname, attribute,'A1:V1',1);
for i = 2 :length(sheetnames)
    if i <= 4 
         xlswrite(outputname, attribute_redox,i,'A1:G1');
    else
         xlswrite(outputname, attribute_raw,i,'A1:E1');
    end
    
end

    disp('start writing ');
%end
%folder是整理過的動物資料編號
datefiles = dir(folder_name);

%mean sigma SE 之column編號
col_const = struct('mean', 2, 'sigma', 10 ,'se' , 14);
%L R_normalizedL   R_ori 在excel檔案裏面第一次出現的時候
row_const = struct('L_335', 3 ,'L_460', 4,'L_redox_nor', 5 ,'L_redox_ori', 6 ,...
    'R_335', 7, 'R_460', 8, 'R_redox_nor', 9, 'R_redox_ori',10);
% for adding value each round
row_names = fieldnames(row_const);
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
           filename = strcat(currentdate,'\',imagefiles(j,1).name, '\','result_new.xls'); 
           if(exist(filename,'file') ==2 )
               %read the 'record' sheet
               [num,txt,raw] = xlsread(filename,1);
               xlsappend(outputname, raw, 1 );
               disp(imagefiles(j,1).name);
           end
            
        end
    end
    if(datefiles(i,1).isdir)
         sheetsappend(datefiles(i,1).name, outputname, row_const, col_const, length(sheetnames) );
    end
     %the difference of each row is 9
     for j =1: length(row_names)
         row_const.(row_names{j,1}) = row_const.(row_names{j,1})+ 9 ;
     end
    
     
end
disp('finished');
end

