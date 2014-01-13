folder_name = 'D:\Dropbox\ROI_avgIntensity_RedoxImg_SD_SE_20131218\Combine_result\12';
outputname = strcat(folder_name, '\combined_result.xls');

attribute = {'date' 'avgIntensity_gray' 'avgIntensity_red' 'avgIntensity_green' 'avgIntensity_blue'};
if(exist(outputname,'file') == 0)
    xlswrite(outputname,attribute );
     disp('yo');
end
%folder是整理過的動物資料編號
datefiles = dir(folder_name);
%date是其日期

for i = 3 : length(datefiles)
    currentdate = strcat(folder_name , '\', datefiles(i,1).name);
    %當日底下的資料
    disp(datefiles(i,1).name);
    if(datefiles(i,1).isdir)
        xlsappend(outputname, {datefiles(i,1).name});
    end
    imagefiles = dir(currentdate);
    for j = 3:length(imagefiles)
        if (imagefiles(j,1).isdir)
           filename = strcat(currentdate,'\',imagefiles(j,1).name, '\','result.xls');
            
           if(exist(filename,'file') ==2 )
               [num,txt,raw] = xlsread(filename);
               xlsappend(outputname, raw);
               
              % disp(filename)
              
               disp(imagefiles(j,1).name)
           end
            
        end
    end
    
end