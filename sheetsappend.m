function []= sheetsappend(dayname, outputname, row_const, col_const, sheetnum )
    %read value from the sorted xls file
    [num,txt,raw] = xlsread(outputname,1);
    col_names = fieldnames(col_const);

    % difference with sheetnum and col_names
    % because col names only 1~3 (mean, sigma, SE)
    % ex : col_names{i - 1, 1} because the "i" is for sheetnum
      diff =[1 ,4];
   
    % i = 2 because it has title on the first row
    for i = 2 : sheetnum
        if(i <=4 )
        % temp_num is the original L,R redox ratio and normalized L,R redox
        % I put it into an array to make the code shorter
        % ratio_ori and ratio_nor just the R/L 
             temp_col_attri = col_const.(col_names{i - diff(1), 1});
             temp_num = [raw(row_const.L_redox_ori, temp_col_attri),...
             raw(row_const.R_redox_ori, temp_col_attri), ...
             raw(row_const.L_redox_nor, temp_col_attri),...
             raw(row_const.R_redox_nor, temp_col_attri) ];
         
             ratio_ori =(num(row_const.R_redox_ori-1,temp_col_attri) /num(row_const.L_redox_ori-1,temp_col_attri));
             ratio_nor =(num(row_const.R_redox_nor-1,temp_col_attri) /num(row_const.L_redox_nor-1,temp_col_attri));
             temp_write = [dayname,temp_num,ratio_ori, ratio_nor];
             xlsappend(outputname, temp_write, i );
        else
             temp_col_attri = col_const.(col_names{i - diff(2), 1});
             temp_num = [raw(row_const.L_335, temp_col_attri),...
             raw(row_const.R_335, temp_col_attri), ...
             raw(row_const.L_460, temp_col_attri),...
             raw(row_const.R_460, temp_col_attri) ];
             temp_write = [dayname,temp_num];
             xlsappend(outputname, temp_write, i );
        end

    end
   
end

