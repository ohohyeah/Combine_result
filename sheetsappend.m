function []= sheetsappend(dayname, outputname, row_const, col_const, sheetnum )
    %read value from the sorted xls file
    [num,txt,raw] = xlsread(outputname,1);
    col_names = fieldnames(col_const);
for i = 2 : sheetnum
    %temp_num is the original L,R redox ratio and normalized L,R redox
    % I put it into an array to make the code shorter
    % ratio_ori and ratio_nor just the R/L 
     temp_num = [raw(row_const.L_redox_ori,col_const.(col_names{i - 1, 1})), raw(row_const.R_redox_ori,col_const.(col_names{i - 1, 1})) raw(row_const.L_redox_nor,col_const.(col_names{i - 1, 1})), raw(row_const.R_redox_nor,col_const.(col_names{i - 1, 1})) ];
     ratio_ori =(num(row_const.R_redox_ori-1,col_const.(col_names{i - 1, 1})) /num(row_const.L_redox_ori-1,col_const.(col_names{i - 1, 1})));
     ratio_nor =(num(row_const.R_redox_nor-1,col_const.(col_names{i - 1, 1})) /num(row_const.L_redox_nor-1,col_const.(col_names{i - 1, 1})));
     temp_write = [dayname,temp_num,ratio_ori, ratio_nor];
     xlsappend(outputname, temp_write, i );
end
    
end

