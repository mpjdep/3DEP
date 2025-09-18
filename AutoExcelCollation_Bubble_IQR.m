% Searches a folder for all the raw 3DEP results. Collates all repeats into a
% single excel sheet and auto-generates the average/sd table.
% IMPORTANT:
% Will treat all samples in a folder as the same sample but different
% technical repeats. Seperate different samples (or treatments) into
% different folders first.
%
% OUTPUT:
%   \Full_Data_Filtered.xlsx: Contains all the datapoints from all technical
%   repeats, a table of frequencies with the average polarisability value
%   and stDev per frequency. The excel sheet is dynamic so anomalies can be
%   removed by manually removing them from column'B'.
%   Datapoints identified by the 3DEP as a 'Bubble', or identified to be
%   outside the IQR of the frequency group are automatically removed.
%
% Authors: MP Johnson, AA Hamka
% Release: 2025
% https://github.com/mpjdep/3DEP
%
% First publication to utilise this code:


folderPath = uigetdir('C:\', 'Select Folder');
if folderPath == 0
    disp('No folder selected. Exiting script.');
    return;
end
disp('Folder selected');
csvFiles = dir(fullfile(folderPath, '*.csv'));
numFiles = length(csvFiles);

if numFiles == 0
    disp('No CSV files found in the selected folder.');
    return;
end
%% Data import and formatting
startRow = 1;
rowOffset = 20; % Number of rows per CSV file
fullData = [];
% Loop through each CSV file
disp('Loading and formatting repeats...');
for i = 1:numFiles
    csvFileName = fullfile(folderPath, csvFiles(i).name);
    csvData = readtable(csvFileName, 'VariableNamingRule', 'preserve',  'EmptyValue',0);
    if size(csvData, 1) >= 20
        extractedData = csvData;
        [~, fileName, ~] = fileparts(csvFiles(i).name);
        lastChar = fileName(end);
        lastCharCell = cell(size(extractedData, 1), 1);
        [lastCharCell{:}] = deal(lastChar);
        if isnumeric(extractedData{:, 3})
            twoColData = extractedData(:,1:2);
            extractedData = [twoColData, lastCharCell];
        else
            extractedData{:, 3} = strcat(lastCharCell, extractedData{:, 3});  
        end
        extractedDataCell = table2cell(extractedData);
        fullData = [fullData; extractedDataCell];
    else
        disp(['Skipping file ', csvFiles(i).name, ' as it does not have enough rows.']);
    end
end
disp(['Data compiled from ', num2str(numFiles), ' repeats']);
        sortedData = sortrows(fullData, 1);

%% Remove IQR outliers
IQRdata = [];
sortedData(:, 4) = {[]};
for i = 1:length(sortedData)
   if contains(sortedData(i,3), 'Bubble')
        sortedData(i,4) = sortedData(i,2);
        sortedData(i,2) = {double.empty(0)};
    end
end

uniqueFreq = unique(cell2mat(sortedData(:, 1))); % Get unique frequencies
groups = cell(size(sortedData, 1), 1); % Initialize groups cell array
for i = 1:numel(uniqueFreq)
    idx = cellfun(@(x) isequal(x, uniqueFreq(i)), sortedData(:, 1));
    groups(idx) = {i};
end

for g = 1:numel(uniqueFreq)
    groupIdx = cell2mat(groups) == g;
    groupData = sortedData(groupIdx, :);
    A = groupData(:, 2);
    B = cell2mat(A);
    C  = groupData(:, 3);
    Q1 = prctile(B, 25);
    Q3 = prctile(B, 75);
    IQR = Q3 - Q1;
    
    % Filter data points outside the IQR and move to column 4
    outliersIdx = (B < Q1 - 1.5 * IQR) | (B > Q3 + 1.5 * IQR);
    groupData(outliersIdx, 4) = groupData(outliersIdx, 2);
    groupData(outliersIdx, 2) = {double.empty(0)}; % Set outliers in column 2 to empty
    sortedData(groupIdx, :) = groupData;
end
%% Excel Export
excelApp = actxserver('Excel.Application');
workbook = excelApp.Workbooks.Add();
excelApp.Visible = 1;
outputFile = fullfile(folderPath, 'Full_Data_Filtered.xlsx');
fullDataTable = array2table(sortedData, 'VariableNames', {'Frequency', 'Intensity', 'Notes', 'RemovedData'});
disp('Data sorted');
writetable(fullDataTable, outputFile);
disp('Data export completed');

%% Average and stDev table generation
disp('Generating table...');
workbook = excelApp.Workbooks.Open(outputFile);
sheet = workbook.Sheets.Item(1); % Assuming you want to work with the first sheet

cellRange = sheet.Range('G1');
cellRange.Formula = 'Frequency, Hz';
cellRange = sheet.Range('H1');
cellRange.Formula = 'Average polarisability, arb. units';
cellRange = sheet.Range('I1');
cellRange.Formula = 'Stdev';
cellRange = sheet.Range('F2');
cellRange.Formula = '2';
cellRange = sheet.Range('F3');
cellRange.Formula = 2 + numFiles;
cellRange = sheet.Range('G2');
cellRange.Formula = '=INDIRECT("A"&F2)'; 
cellRange = sheet.Range('H2');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F2):INDIRECT("B"&(F3-1)))';
cellRange = sheet.Range('I2');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F2):INDIRECT("B"&(F3-1)))';
disp('.');
cellRange = sheet.Range('F4');
cellRange.Formula = 2 + 2*numFiles;
cellRange = sheet.Range('F5');
cellRange.Formula = 2 + 3*numFiles;
cellRange = sheet.Range('F6');
cellRange.Formula = 2 + 4*numFiles;
cellRange = sheet.Range('F7');
cellRange.Formula = 2 + 5*numFiles;
cellRange = sheet.Range('F8');
cellRange.Formula = 2 + 6*numFiles;
cellRange = sheet.Range('F9');
cellRange.Formula = 2 + 7*numFiles;
cellRange = sheet.Range('F10');
cellRange.Formula = 2 + 8*numFiles;
cellRange = sheet.Range('F11');
cellRange.Formula = 2 + 9*numFiles;
cellRange = sheet.Range('F12');
cellRange.Formula = 2 + 10*numFiles;
cellRange = sheet.Range('F13');
cellRange.Formula = 2 + 11*numFiles;
cellRange = sheet.Range('F14');
cellRange.Formula = 2 + 12*numFiles;
cellRange = sheet.Range('F15');
cellRange.Formula = 2 + 13*numFiles;
cellRange = sheet.Range('F16');
cellRange.Formula = 2 + 14*numFiles;
cellRange = sheet.Range('F17');
cellRange.Formula = 2 + 15*numFiles;
cellRange = sheet.Range('F18');
cellRange.Formula = 2 + 16*numFiles;
cellRange = sheet.Range('F19');
cellRange.Formula = 2 + 17*numFiles;
cellRange = sheet.Range('F20');
cellRange.Formula = 2 + 18*numFiles;
cellRange = sheet.Range('F21');
cellRange.Formula = 2 + 19*numFiles;
cellRange = sheet.Range('F22');
cellRange.Formula = 2 + 20*numFiles;
disp('..');
cellRange = sheet.Range('G3');
cellRange.Formula = '=INDIRECT("A"&F3)'; 
cellRange = sheet.Range('G4');
cellRange.Formula = '=INDIRECT("A"&F4)'; 
cellRange = sheet.Range('G5');
cellRange.Formula = '=INDIRECT("A"&F5)'; 
cellRange = sheet.Range('G6');
cellRange.Formula = '=INDIRECT("A"&F6)'; 
cellRange = sheet.Range('G7');
cellRange.Formula = '=INDIRECT("A"&F7)'; 
cellRange = sheet.Range('G8');
cellRange.Formula = '=INDIRECT("A"&F8)'; 
cellRange = sheet.Range('G9');
cellRange.Formula = '=INDIRECT("A"&F9)'; 
cellRange = sheet.Range('G10');
cellRange.Formula = '=INDIRECT("A"&F10)'; 
cellRange = sheet.Range('G11');
cellRange.Formula = '=INDIRECT("A"&F11)'; 
cellRange = sheet.Range('G12');
cellRange.Formula = '=INDIRECT("A"&F12)'; 
cellRange = sheet.Range('G13');
cellRange.Formula = '=INDIRECT("A"&F13)'; 
cellRange = sheet.Range('G14');
cellRange.Formula = '=INDIRECT("A"&F14)'; 
cellRange = sheet.Range('G15');
cellRange.Formula = '=INDIRECT("A"&F15)'; 
cellRange = sheet.Range('G16');
cellRange.Formula = '=INDIRECT("A"&F16)'; 
cellRange = sheet.Range('G17');
cellRange.Formula = '=INDIRECT("A"&F17)'; 
cellRange = sheet.Range('G18');
cellRange.Formula = '=INDIRECT("A"&F18)'; 
cellRange = sheet.Range('G19');
cellRange.Formula = '=INDIRECT("A"&F19)'; 
cellRange = sheet.Range('G20');
cellRange.Formula = '=INDIRECT("A"&F20)'; 
cellRange = sheet.Range('G21');
cellRange.Formula = '=INDIRECT("A"&F21)'; 
disp('...');
cellRange = sheet.Range('H3');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F3):INDIRECT("B"&(F4-1)))';
cellRange = sheet.Range('H4');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F4):INDIRECT("B"&(F5-1)))';
cellRange = sheet.Range('H5');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F5):INDIRECT("B"&(F6-1)))';
cellRange = sheet.Range('H6');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F6):INDIRECT("B"&(F7-1)))';
cellRange = sheet.Range('H7');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F7):INDIRECT("B"&(F8-1)))';
cellRange = sheet.Range('H8');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F8):INDIRECT("B"&(F9-1)))';
cellRange = sheet.Range('H9');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F9):INDIRECT("B"&(F10-1)))';
cellRange = sheet.Range('H10');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F10):INDIRECT("B"&(F11-1)))';
cellRange = sheet.Range('H11');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F11):INDIRECT("B"&(F12-1)))';
cellRange = sheet.Range('H12');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F12):INDIRECT("B"&(F13-1)))';
cellRange = sheet.Range('H13');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F13):INDIRECT("B"&(F14-1)))';
cellRange = sheet.Range('H14');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F14):INDIRECT("B"&(F15-1)))';
cellRange = sheet.Range('H15');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F15):INDIRECT("B"&(F16-1)))';
cellRange = sheet.Range('H16');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F16):INDIRECT("B"&(F17-1)))';
cellRange = sheet.Range('H17');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F17):INDIRECT("B"&(F18-1)))';
cellRange = sheet.Range('H18');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F18):INDIRECT("B"&(F19-1)))';
cellRange = sheet.Range('H19');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F19):INDIRECT("B"&(F20-1)))';
cellRange = sheet.Range('H20');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F20):INDIRECT("B"&(F21-1)))';
cellRange = sheet.Range('H21');
cellRange.Formula = '=AVERAGE(INDIRECT("B"&F21):INDIRECT("B"&(F22-1)))';
disp('....');
cellRange = sheet.Range('I3');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F3):INDIRECT("B"&(F4-1)))';
cellRange = sheet.Range('I4');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F4):INDIRECT("B"&(F5-1)))';
cellRange = sheet.Range('I5');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F5):INDIRECT("B"&(F6-1)))';
cellRange = sheet.Range('I6');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F6):INDIRECT("B"&(F7-1)))';
cellRange = sheet.Range('I7');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F7):INDIRECT("B"&(F8-1)))';
cellRange = sheet.Range('I8');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F8):INDIRECT("B"&(F9-1)))';
cellRange = sheet.Range('I9');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F9):INDIRECT("B"&(F10-1)))';
cellRange = sheet.Range('I10');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F10):INDIRECT("B"&(F11-1)))';
cellRange = sheet.Range('I11');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F11):INDIRECT("B"&(F12-1)))';
cellRange = sheet.Range('I12');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F12):INDIRECT("B"&(F13-1)))';
cellRange = sheet.Range('I13');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F13):INDIRECT("B"&(F14-1)))';
cellRange = sheet.Range('I14');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F14):INDIRECT("B"&(F15-1)))';
cellRange = sheet.Range('I15');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F15):INDIRECT("B"&(F16-1)))';
cellRange = sheet.Range('I16');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F16):INDIRECT("B"&(F17-1)))';
cellRange = sheet.Range('I17');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F17):INDIRECT("B"&(F18-1)))';
cellRange = sheet.Range('I18');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F18):INDIRECT("B"&(F19-1)))';
cellRange = sheet.Range('I19');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F19):INDIRECT("B"&(F20-1)))';
cellRange = sheet.Range('I20');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F20):INDIRECT("B"&(F21-1)))';
cellRange = sheet.Range('I21');
cellRange.Formula = '=STDEV.P(INDIRECT("B"&F21):INDIRECT("B"&(F22-1)))';
disp('Table finished');
%% Excel cleanup
workbook.Save();
disp('Process Complete');
%workbook.Close(false); % false simply avoids another save on closing
excelApp.Quit(); % closes Excel the application
excelApp.delete(); % deletes the Excel COM object behind the scenes consuming memory
disp('Excel processes stopped');
disp(outputFile);