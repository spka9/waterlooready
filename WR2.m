[~,SheetNames]  = xlsfinfo('attaug29.xlsx'); % pull the names of all sheets into a 1x28 cell array
nSheets = length(SheetNames);
Data = ones(32,16,nSheets)*99; %communities with less than 32 students will have each excess row filled with 99s
for ii=1:nSheets
    sheetData = readmatrix('attaug29.xlsx', 'Sheet', ii); %reads values of iith sheet and assigns to temp var sheetData
    attendance = sheetData(6:end, 3:end-2);%take the portion of each sheet that contains attendance and assign it to temp var attendance
    attendance(isnan(attendance)) = 0; %wherever attendance is blank write a 0
    Data(1:size(attendance, 1),1:size(attendance, 2),ii) = attendance; %rewrite the iith sheet of Data to be attendance
end


header = [];
for ii=1:nSheets
[~, ~, raw] = xlsread('attaug29.xlsx',ii);
headertemp = raw(1:4,3:18);
header = cat(3, header, headertemp); %capture the week, day, and time of each meeting in a cell array
end

days_of_week = {'Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday', 'Saturday', 'Sunday'};
ideal_data_info=zeros(5, 16, nSheets);
%create a double matrix capturing header data
for ss=1:nSheets
    for ii=1:16
    ideal_data_info(1, 1:2:16, ss) = 1:8; %make top 2 rows where week number is stored
    ideal_data_info(1, 2:2:16, ss) = 1:8;
    ideal_data_info(2, 1:2:16, ss) = 1; %meeting number within the week
    ideal_data_info(2, 2:2:16, ss) = 2;
        if isnan(header{3, ii, ss}) %if the day of the week is blank do nothing
        else
            ideal_data_info(5, ii, ss) = 1; %row 5 captures if there is data recorded or not
            for cc=1:7
                if ismember(header{3, ii, ss}, days_of_week{cc})
                    ideal_data_info(3, ii, ss) = cc; %row 3 captures the day of the week of the meeting
                end
            end
            ideal_data_info(4, ii, ss) = floor((header{4, ii, ss})*24); %row 4 is the hour of the meeting
        end
    end
end

ideal_data_info_labels = ["weeknumb"; "meetnumofweek"; "dayofweek"; "hour"; "istheredata"];

%RESEARCH QUESTION 1
%MOST POPULAR DAY OF THE WEEK
vec_mon = [];
vec_tues = [];
vec_wed = [];
vec_thur = [];
vec_fri = [];
vec_sat = [];
vec_sun = [];

dayofweek_oa = zeros(4, 7);


stdthurs = [];
stdsun = [];


for ss=1:nSheets
    for ii=1:16
        if ideal_data_info(3, ii, ss) == 1
            vec_mon = [vec_mon; Data(:, ii, ss)];
            dayofweek_oa(1, 1) = dayofweek_oa(1, 1) + 1;
        elseif ideal_data_info(3, ii, ss) == 2
            dayofweek_oa(1, 2) = dayofweek_oa(1, 2) + 1;
            vec_tues = [vec_tues; Data(:, ii, ss)];
        elseif ideal_data_info(3, ii, ss) == 3
            vec_wed = [vec_wed; Data(:, ii, ss)];
            dayofweek_oa(1, 3) = dayofweek_oa(1, 3) + 1;
        elseif ideal_data_info(3, ii, ss) == 4
            vec_thur = [vec_thur; Data(:, ii, ss)];
            dayofweek_oa(1, 4) = dayofweek_oa(1, 4) + 1;
            no99 = Data(:, ii, ss);
            no99(no99 == 99) = [];
            no99n = numel(no99);
            stdthurs = [stdthurs (sum(no99))/no99n];
            
        elseif ideal_data_info(3, ii, ss) == 5
            vec_fri = [vec_fri; Data(:, ii, ss)];
            dayofweek_oa(1, 5) = dayofweek_oa(1, 5) + 1;
        elseif ideal_data_info(3, ii, ss) == 6
            vec_sat = [vec_sat; Data(:, ii, ss)];
            dayofweek_oa(1, 6) = dayofweek_oa(1, 6) + 1;
        elseif ideal_data_info(3, ii, ss) == 7
            vec_sun = [vec_sun; Data(:, ii, ss)];
            dayofweek_oa(1, 7) = dayofweek_oa(1, 7) + 1;
            no99 = Data(:, ii, ss);
            no99(no99 == 99) = [];
            no99n = numel(no99);
            stdsun = [stdsun (sum(no99))/no99n];
        end
    end
end

stdsun2 = std(stdsun)
stdthurs2 = std(stdthurs)

vec_mon(vec_mon == 99) = [];
vec_tues(vec_tues == 99) = [];
vec_wed(vec_wed == 99) = [];
vec_thur(vec_thur == 99) = [];
vec_fri(vec_fri == 99) = [];
vec_sat(vec_sat == 99) = [];
vec_sun(vec_sun == 99) = [];

dayofweek_oa(2, 1) = (numel(vec_mon));   
dayofweek_oa(2, 2) = (numel(vec_tues));  
dayofweek_oa(2, 3) = (numel(vec_wed));
dayofweek_oa(2, 4) = (numel(vec_thur));   
dayofweek_oa(2, 5) = (numel(vec_fri));  
dayofweek_oa(2, 6) = (numel(vec_sat));  
dayofweek_oa(2, 7) = (numel(vec_sun));  

dayofweek_oa(3, 1) = sum(vec_mon);   
dayofweek_oa(3, 2) = sum(vec_tues);  
dayofweek_oa(3, 3) = sum(vec_wed);
dayofweek_oa(3, 4) = sum(vec_thur);   
dayofweek_oa(3, 5) = sum(vec_fri);  
dayofweek_oa(3, 6) = sum(vec_sat);  
dayofweek_oa(3, 7) = sum(vec_sun);  

dayofweek_oa(4, 1) = dayofweek_oa(3, 1)/dayofweek_oa(2, 1);
dayofweek_oa(4, 2) = dayofweek_oa(3, 2)/dayofweek_oa(2, 2);
dayofweek_oa(4, 3) = dayofweek_oa(3, 3)/dayofweek_oa(2, 3);
dayofweek_oa(4, 4) = dayofweek_oa(3, 4)/dayofweek_oa(2, 4);
dayofweek_oa(4, 5) = dayofweek_oa(3, 5)/dayofweek_oa(2, 5);
dayofweek_oa(4, 6) = dayofweek_oa(3, 6)/dayofweek_oa(2, 6);
dayofweek_oa(4, 7) = dayofweek_oa(3, 7)/dayofweek_oa(2, 7); 

plot(dayofweek_oa(4, :))

%RESEARCH QUESTION 2
%MOST POPULAR HOUR

time = zeros(4, 24);

hr11 = [];
hr14 = [];
hr17 = [];
hr18 = [];
hr20 = [];



for ss=1:nSheets
    for ii=1:16
        hour = ideal_data_info(4, ii, ss);
        if hour == 0
        else
        time(1, hour) = time(1, hour) +1;
        temp = Data(:, ii, ss);
        temp(temp == 99) = [];
        time(2, hour) = time(2, hour) + numel(temp);
        time(3, hour) = time(3, hour) + sum(temp);
            if hour == 11
                hr11 = [hr11 sum(temp)/numel(temp)];
            elseif hour == 14
                hr14 = [hr14 sum(temp)/numel(temp)];
            elseif hour == 17
                hr17 = [hr17 sum(temp)/numel(temp)];
            elseif hour == 18
                hr18 = [hr18 sum(temp)/numel(temp)];
            elseif hour == 20
                hr20 = [hr20 sum(temp)/numel(temp)];
            end
        end
    end
end

stdhr = zeros(1, 24);
stdhr(11) = std(hr11);
stdhr(14) = std(hr14);
stdhr(17) = std(hr17);
stdhr(18) = std(hr18);
stdhr(20) = std(hr20);


time(4, :) = time(3, :)./time(2, :);

plot(time(4, :))

% RESEARCH QUESTION 3
% DIVERSE TIMES

diverse_time_table = zeros(7, 2);
prevweek = [];
stddev_div = [];
stddev_notdiv = [];
stddev_div_erp = [];
stddev_notdiv_erp = [];

for ss=1:nSheets
    prevweek = []
    for ii=1:2:13
        hour1 = ideal_data_info(4, ii, ss);
        hour2 = ideal_data_info(4, ii+1, ss);
        if hour1~=0 && hour2~=0
            if abs(hour1-hour2)>3
                div_notdiv = 1;
            else
                div_notdiv = 2;
            end
            temp = Data(:, ii:ii+1, ss);
            diverse_time_table(1, div_notdiv) = diverse_time_table(1, div_notdiv)+1;
            temp(temp == 99) = [];
            numel(temp)
            diverse_time_table(2, div_notdiv) = diverse_time_table(2, div_notdiv) + numel(temp);
            diverse_time_table(3, div_notdiv) = diverse_time_table(3, div_notdiv) + sum(temp, 'all');
            if div_notdiv == 1
                stddev_div = [stddev_div (sum(temp(:, 1)))/(numel(temp)/2) (sum(temp(:, 2)))/(numel(temp)/2)];
            else
               stddev_notdiv = [stddev_notdiv (sum(temp(:, 1)))/(numel(temp)/2) (sum(temp(:, 2)))/(numel(temp)/2)];
            end
            
            temp = reshape(temp, [], 2);
            %erp
            
            temp = sum(temp, 2);
            temp(temp == 2) = 1;

            if ii == 1
                break
            else
                erpcomp = [temp prevweek];
                erpcomp = sum(sum(erpcomp, 2) == 2);

                diverse_time_table(6, div_notdiv) = diverse_time_table(6, div_notdiv) + erpcomp;
                if div_notdiv == 1
                    stddev_div_erp = [stddev_div_erp erpcomp/numel(temp)];
                else
                    stddev_notdiv_erp = [stddev_div_erp erpcomp/numel(temp)];
                end
            end
            prevweek = temp;
        else
        end
    end
    diverse_time_table(5, div_notdiv) = diverse_time_table(5, div_notdiv) + numel(temp(:, 1))*((sum(sum(reshape(ideal_data_info(5, :, ss), 2, []))>0))-1);
end
% 
diverse_time_table(4, :) = diverse_time_table(3, :)./diverse_time_table(2, :);
diverse_time_table(7, :) = diverse_time_table(6, :)./diverse_time_table(5, :);
diverse_time_table(8, :) = [std(stddev_div) std(stddev_notdiv)];
diverse_time_table(9, :) = [std(stddev_div_erp) std(stddev_notdiv_erp)];


%RESEARCH QUESTION 4 consistent times
consistent = zeros(6, 2);
std4cons = [];
std4div = [];
std4consret = [];
std4divret = [];

for ss=1:nSheets
    for ii=3:16
        current = ideal_data_info(3:4, ii, ss);
        prev = ideal_data_info(3:4, ii-2, ss);
        if prev == [0; 0]
            for cc = ii-1:-2:0
                prev = ideal_data_info(3:4, ii-cc, ss);
                if prev ~= [0; 0]
                    break
                end
            end
        end
        
        if current(1)~= 0 && prev(1)~=0
            if current == prev
                cons=1;
            else
                cons=2;
            end
        consistent(1, cons) = consistent(1, cons) + 1;
        temp = [Data(:, ii, ss) Data(:, ii-2, ss)];
        temp(temp == 99) = [];
        temp = reshape(temp, [], 2);
        consistent(2, cons) = consistent(2, cons) + size(temp, 1);
        consistent(3, cons) = consistent(3, cons) + sum(temp(:, 1));
        if cons == 1
            std4cons = [std4cons sum(temp(:, 1))/size(temp, 1)];
        else
            std4div = [std4div sum(temp(:, 1))/size(temp, 1)];
        end
        temp = sum(temp, 2);
        temp = temp == 2;
        consistent(6, cons) = consistent(6, cons) + sum(temp);
        if cons == 1
            std4consret = [std4consret sum(temp)/size(temp, 1)];
        else
            std4divret = [std4divret sum(temp)/size(temp, 1)];
        end
        end
    end
end

consistent(5, :) = consistent(2, :) - consistent(1, :);
consistent(4, :) = consistent(3, :)./consistent(2, :);
consistent(7, :) = consistent(6, :)./consistent(5, :);
consistent(8, 1) = std(std4cons);
consistent(8, 2) = std(std4div);
consistent(9, 1) = std(std4consret);
consistent(9, 2) = std(std4divret);


studentraw = readtable('studentinfo.xlsx');
acell = cell(827, 1);
studentraw = [studentraw acell];
student = zeros(827, 9);
student(:, 1) = 1:827;
student(:, 2:3) = studentraw{:, 2:3};
PeerDemRaw=readtable('PeerDem.xlsx');

for ii = 1:28
    studentraw{(studentraw{:, 2} == PeerDemRaw{ii, 6}), 7} = PeerDemRaw{ii, 7};
end


stream_table = zeros(8, 2);
academiclevel = zeros(8, 5);
coop_aca = zeros(8, 2);
stddevstream_diff = [];
stddevstream_same = [];
stddevstream_diff_erp = [];
stddevstream_same_erp = [];
stddevcoop = [];
stddevaca = [];
stddevcoop_erp = [];
stddevaca_erp = [];
std2A = [];
std2B = [];
std3A = [];
std3B = [];
std4A = [];
std2Aerp = [];
std2Berp = [];
std3Aerp = [];
std3Berp = [];
std4Aerp = [];


for ss=1:nSheets
    cons = PeerDemRaw{ss, 16} +1;
    if isequal(PeerDemRaw{ss, 15}, {'2A'})
        conslevel=1;
    elseif isequal(PeerDemRaw{ss, 15}, {'2B'})
        conslevel=2;
    elseif isequal(PeerDemRaw{ss, 15}, {'3A'})
        conslevel=3;
    elseif isequal(PeerDemRaw{ss, 15}, {'3B'})
        conslevel=4;
    elseif isequal(PeerDemRaw{ss, 15}, {'4A'})
        conslevel=5;
    end
    if isequal(PeerDemRaw{ss, 14}, {'co-op'})
        consco=1;
    elseif isequal(PeerDemRaw{ss, 14}, {'academic'})
        consco=2;
    end  
    temp = Data(:, :, ss);
    rowremove = find(temp(:, 1) == 99);
    temp(rowremove, :) = [];
%     row 1 is number of attendance opps
%     row 2 is number of attendances
    numofstu = numel(temp(:, 1));

   
    academiclevel(1, conslevel) = academiclevel(1, conslevel) + sum(ideal_data_info(5, :, ss))*numel(temp(:, 1));
    academiclevel(2, conslevel) = academiclevel(2, conslevel) + sum(temp, 'all');
    academiclevel(7, conslevel) = academiclevel(7, conslevel) + sum(ideal_data_info(5, :, ss));
    
    coop_aca(1, consco) = coop_aca(1, consco) + sum(ideal_data_info(5, :, ss))*numel(temp(:, 1));
    coop_aca(2, consco) = coop_aca(2, consco) + sum(temp, 'all');
    coop_aca(7, consco) = coop_aca(7, consco) + sum(ideal_data_info(5, :, ss));
    
    tempstream = sum(temp)./numel(temp(:, 1));
            
    if consco == 1
        stddevcoop = [stddevcoop tempstream];
    else
        stddevaca = [stddevaca tempstream];
    end
    
     if conslevel == 1
        std2A = [std2A tempstream];
     elseif conslevel == 2
        std2B = [std2B tempstream];
     elseif conslevel == 3
        std3A = [std3A tempstream];
     elseif conslevel == 4
        std3B = [std3B tempstream];
     else
        std4A = [std4A tempstream];
    end
    
    for jj = 1:2:15
        if sum(ideal_data_info(5, jj:jj+1, ss)) == 0
            temp(:,jj:jj+1) = 99;
        end
    end
    colremove = find(temp(1, :) == 99);
    temp(:, colremove) = [];
    if (sum(sum((reshape(ideal_data_info(5, :, ss), 2, []))>0)))>0
        academiclevel(4, conslevel) = academiclevel(4, conslevel) + numel(temp(:, 1))*((sum(sum(reshape(ideal_data_info(5, :, ss), 2, []))>0))-1);
        coop_aca(4, consco) = coop_aca(4, consco) + numel(temp(:, 1))*((sum(sum(reshape(ideal_data_info(5, :, ss), 2, []))>0))-1);
        
        coop_aca(8, consco) = coop_aca(8, consco) + ((sum(sum(reshape(ideal_data_info(5, :, ss), 2, []))>0))-1);
        academiclevel(8, conslevel) = academiclevel(8, conslevel) + ((sum(sum(reshape(ideal_data_info(5, :, ss), 2, []))>0))-1);

        
        for ii = 1:2:(numel(temp(1,:)) - 3)    
            temp_wk1 = sum(temp(:, ii:ii+1), 2);
            temp_wk2 = sum(temp(:, ii+2:ii+3), 2);
            temp_wk = [temp_wk1 temp_wk2];
            temp_wk(temp_wk == 2) = 1;
            temp_wk = sum(temp_wk, 2);
            temp_wk = temp_wk == 2;
            temp_wk = sum(temp_wk);
            academiclevel(5, conslevel) = academiclevel(5, conslevel) + temp_wk;
            coop_aca(5, consco) = coop_aca(5, consco) + temp_wk;
            
            tempstreamerp = temp_wk/numofstu;

                if consco == 1
                    stddevcoop_erp = [stddevcoop_erp tempstreamerp];
                else
                    stddevaca_erp = [stddevaca_erp tempstreamerp];
                end
                
    
                 if conslevel == 1
                    std2Aerp = [std2Aerp tempstreamerp];
                 elseif conslevel == 2
                    std2Berp = [std2Berp tempstreamerp];
                 elseif conslevel == 3
                    std3Aerp = [std3Aerp tempstreamerp];
                 elseif conslevel == 4
                    std3Berp = [std3Berp tempstreamerp];
                 else
                    std4Aerp = [std4Aerp tempstreamerp];
                end
        end
    end
end
 

academiclevel(3, :) = academiclevel(2, :)./academiclevel(1, :);
academiclevel(6, :) = academiclevel(5, :)./academiclevel(4, :);
plot(1:5, academiclevel(3, :), 1:5, academiclevel(6, :))

coop_aca(3, :) = coop_aca(2, :)./coop_aca(1, :);
coop_aca(6, :) = coop_aca(5, :)./coop_aca(4, :);

stddevcoop2 = std(stddevstream_diff);
stddevaca2 = std(stddevstream_same);
stddevcooperp = std(stddevcoop_erp);
stddevacaerp = std(stddevaca_erp);

stdlevel = zeros(2,5);
stdlevel(1,1) = std(std2A);
stdlevel(1,2) = std(std2B);
stdlevel(1,3) = std(std3A);
stdlevel(1,4) = std(std3B);
stdlevel(1,5) = std(std4A);
stdlevel(2,1) = std(std2Aerp);
stdlevel(2,2) = std(std2Berp);
stdlevel(2,3) = std(std3Aerp);
stdlevel(2,4) = std(std3Berp);
stdlevel(2,5) = std(std4Aerp);

clearvars -except academiclevel  consistent coop_aca dayofweek_oa  diverse_time_table   time
