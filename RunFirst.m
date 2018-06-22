%% Hello and welcome to the sled lab data analysis script. 
% You can use the setup section to configure your analysis and run this
% script in matlab. It will automatically produce the results you want.

clearvars;  % Remove previous variables
%% Setup section
% Change this to the file path you wish to analyse
fileName = 'C:\Users\n.kent\Documents\ChestClipStudy\CCA - 00030\CCA - 00030_FILTERED.csv';

% Debug files
%fileName = 'C:\Users\n.kent\Documents\EDA_TestFile.csv';
%fileName = 'C:\Users\n.kent\Documents\CCA - 00002_FILTERED.csv';

% Channels - add your channels here
Chanl =  'SledG';
Chan2 =  'HeadGx';
Chan3 =  'HeadGy';
Chan4 =  'HeadGz';
Chan5 =  'NeckFx';
Chan6 =  'NeckFy';
Chan7 =  'NeckMy';
Chan8 =  'NotUsed';
Chan9 =  'NotUsed';
Chanl0 = 'NotUsed';
Chan11 = 'NotUsed';
Chanl2 = 'NotUsed';
Chanl3 = 'NotUsed';
Chanl4 = 'NotUsed';
Chanl5 = 'NotUsed';
Chanl6 = 'NotUsed';
Chanl7 = 'NotUsed';
Chanl8 = 'NotUsed';
% This is the limit to our DTS. If more channels are needed the
% DataAnalysis file will need to be modded.

% Set the limits (If you dont know this yet just leave the current values)
% These values are in milliseconds from start of data
setLowerLimit = 500;
setUpperLimit = 750;

% Now set what you want to calcuate and diplay. Type 'y' or 'n' for the
% things you want to do

createPlots = 'y';
CalculateResultantHeadAcceleration = 'n';
CalculateResultantNeckForce = 'n';
CalculateResultantNeckMoment = 'n';
CalculateHIC = 'n';
CalculateNIC = 'n';
CalculateMaxValues = 'n';
CalculateMinValues = 'n';
CreateReporttxt = 'n';
CreateReportWord = 'n';

% Runtime choices
showPlots = 1;          % 0 to turn off graphs, 1 to show graphs 
smoothGraphs = 1;       % 0 for raw data, 1 for smoothed graphs 
smoothChoice = 5;       % 0:'moving', 1:'lowess', 2:'loess', 3:'sgolay' 4:'rlowess' 5:'rloess'
smoothSpan = 0.05;      % Determines smoothness - must be odd number
timeConversion = 1000;  % Conversion factor for time scale
picSavePath = 'C:\Users\n.kent\Documents\ChestClipStudy'; % Set you save file path here
picFileExt = '.jpeg';   % Set the save file extension - nick(make this automatic)
picResolution = '-r300';% Set save file resolution
picFormatType = '-djpeg';% Set format type - e.g. jpeg, png
printGraphs = 'true';   % 'true' to save graphs file
graphleft = 300;        % graphs position from left of screen
graphBottom = 50;       % graphs position from bottom of screen
graphWidth = 1400;      % graph width
graphHeight = 900;      % graph height

% End of user section - You may now run the script

%% Bug report
% Please place any bugs you find here and I can work them out
%-

%% Ok here we go. This section does all the work 

% @@@@@@@@@@@@@@    WARNING WARNING WARNING    @@@@@@@@@@@@@@@@@@@@@@@@
% ****(Do not change anything below this line or you will break it)****

% Channel array
channels = [string(Chanl); string(Chan2); string(Chan3); string(Chan4);
            string(Chan5); string(Chan6); string(Chan7); string(Chan8);
            string(Chan9); string(Chanl0); string(Chan11); string(Chanl2);
            string(Chanl3); string(Chanl4); string(Chanl5); string(Chanl6);
            string(Chanl7); string(Chanl8)];
        
% Store runtime values
runTimeVals = [ showPlots, smoothGraphs, smoothChoice, smoothSpan,...
                timeConversion];
            
% Store the graph printing properties          
plotVars = {picSavePath, picFileExt, picResolution, picFormatType, printGraphs};
plotSizing = [graphleft graphBottom graphWidth graphHeight];

% Set up object
a = DataAnalysis(fileName, runTimeVals, channels, plotVars, plotSizing);

% Seting limits
a.changeLimits(setLowerLimit, setUpperLimit);

% Creating plots
if strcmp(createPlots ,'yes') || strcmp(createPlots ,'y')
    a.createPlots();
end

% Calculaing resultant head acceleration
if strcmp(CalculateResultantHeadAcceleration, 'yes') || strcmp(CalculateResultantHeadAcceleration, 'y')
    a.resultantHeadAcel();
end

% Calculating resultant neck forces
if strcmp(CalculateResultantNeckForce, 'yes') || strcmp(CalculateResultantNeckForce, 'y')
    a.resultantNeckForce();
end

% Calculating resultant neck moment
if strcmp(CalculateResultantNeckMoment, 'yes') || strcmp(CalculateResultantNeckMoment, 'y')
    a.resultantNeckMoment();
end

% Calculting HIC
if strcmp(CalculateHIC, 'yes') || strcmp(CalculateHIC, 'y')
    a.calcHIC();
end

% Caluclting NIC
if strcmp(CalculateNIC, 'yes') || strcmp(CalculateNIC, 'y')
    a.calcNeckInjuryCriteria();
end

% Finding maximum vaules
if strcmp(CalculateMaxValues, 'yes') || strcmp(CalculateMaxValues, 'y')
    a.calcMaxVals();
end

% Finding minimum values
if strcmp(CalculateMinValues, 'yes') || strcmp(CalculateMinValues, 'y')
    a.calcMinVals();
end

% Create a report - txt
if strcmp(CreateReporttxt, 'yes') || strcmp(CreateReporttxt, 'y')
    a.CreateReporttxt();
end

% Create a report - Word
if strcmp(CreateReportWord, 'yes') || strcmp(CreateReportWord, 'y')
    a.CreateReportWord();
end