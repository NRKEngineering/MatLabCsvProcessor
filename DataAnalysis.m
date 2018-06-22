%% This script will read the data file produced by the DTS data aquisition system
%  and produce the required result as deterimned by the RunFirst script.

classdef DataAnalysis < handle
    % This class allows analysis of the data obtained from the DTS data
    % aquisition system.
    
    % Functions available
    
    % - DataAnalysis(fileName)  // Run first, sets up object
    % - createPlots()           // Creates plots of the data
    % - resultantHeadAcel()     // Calcultes the resultant head acceleration
    % - resultantNeckForce()    // Calcultes the resultant neck force
    % - resultantNeckMoment()   // Calcultes the resultant neck moment
    % - calcHIC()               // Calculates Head Injury Criteria (HIC)
    % - calcNeckInjuryCriteria()// Calculates Neck Injury Criteria (NIC)
    % - calcMaxVals()           // Returns the maximum values
    % - calcMinVals()           // Returns the minimum values
    % - changeLimits(front, rear) // Alows changes to graph limits (in ms)
    
    %% 
    properties(Access = public)
        rawData;
        initVal;
        endVal;
        TestDate;
        TestTime;
        TestId;
        TestDesc;
        sampleRate;
        channelDesc;
        dataChannel;
        SensorSerial;
        engineeringUnit;
        sampleFreq;
        sampleTime;
        data;
        rows;
        columns;
        runTimeVals;
        maxVals;
        minVals;
        HICVal15;
        HICVal36;
        NICVal;
        
        % Channels - set all to off(user will turn on required channels)
        notUsed = -1;
        time =    -1;
        sledG =   -1;
        headGx =  -1;
        headGy =  -1;
        headGz =  -1;
        chestGx = -1;
        chestGy = -1;
        chestGz = -1;
        neckFx =  -1;
        neckFy =  -1;
        neckFz =  -1;
        neckMx =  -1;
        neckMy =  -1;
        neckMz =  -1;
        iliacLFx = -1;
        iliacLMy = -1;
        iliacRFx = -1;
        iliacRMy = -1;
        % Add more channels as nessesary
    
        % Display limits
        cutStart;           % Adjust this to trim front of graph
        cutEnd;            % Adjust this to trim rear of graph
        
        % Save variables
%         plotControls;
         picSavePath;
%         picFileName;
%         picFileExt;
%         picResolution;
%         picFormatType;
%         printGraphs;
        plotVars;
        plotSizing;
    
    end
    %%
    methods
        %% Constructor
        function obj = DataAnalysis(fileName, runTimeVals, channels, plotVars, plotSizing)
            % Store run time variables
            obj.runTimeVals = runTimeVals;
            obj.plotVars = plotVars;
            obj.plotSizing = plotSizing;
            
            % Assign user channel choice
            chanColumns  = size(channels);
            for i = 2:chanColumns
                switch channels(i)
                    case 'SledG'
                       obj.sledG =  i;
                    case 'HeadGx'
                       obj.headGx = i;
                    case 'HeadGy'
                       obj.headGy = i;
                    case 'HeadGz'
                       obj.headGz = i;
                    case 'ChestGx'
                       obj.chestGx = i;
                    case 'ChestGy'
                       obj.chestGy = i;
                    case 'ChestGz'
                       obj.chestGz = i;
                    case 'NeckFx'
                       obj.neckFx = i;
                    case 'NeckFy'
                       obj.neckFy = i;
                    case 'NeckFz'
                       obj.neckFz = i;
                    case 'NeckMx'
                       obj.neckMx = i;
                    case 'NeckMy'
                       obj.neckMy = i;
                    case 'NeckMz'
                       obj.neckMz = i;
                    case 'IliacLFx'
                       obj.iliacLFx = i;
                    case 'IliacLMy'
                       obj.iliacLMy = i;
                    case 'IliacRFx'
                       obj.iliacRFx = i;
                    case 'IliacRMy'
                       obj.iliacRMy = i;
                    case 'NotUsed'
                end
            end
            
            
            
            % Tell user whats going on
            disp('Analysing data...');
            
            
            
            % Checks for file and stores data, headers in table
            if exist( fileName, 'file')
                disp('Reading file');
                obj.rawData = readtable (fileName);
            else
                error(['File ', fileName, ' not found']);
            end
            
            %Store variables in object
            obj.TestDate = obj.rawData{1,2};
            obj.TestTime = obj.rawData.Var2(2);
            obj.TestId =   obj.rawData.Var2(3);
            obj.TestDesc = obj.rawData.Var2(4);
            obj.sampleRate = [obj.rawData.Var2(5),obj.rawData.Var3(5), ...
                              obj.rawData.Var4(5),obj.rawData.Var5(5), ...
                              obj.rawData.Var6(5),obj.rawData.Var7(5), ...
                              obj.rawData.Var8(5)];
            obj.channelDesc = [obj.rawData.Var2(9),obj.rawData.Var3(9), ...
                               obj.rawData.Var4(9),obj.rawData.Var5(9), ...
                               obj.rawData.Var6(9),obj.rawData.Var7(9), ...
                               obj.rawData.Var8(9)];
            obj.dataChannel = [obj.rawData.Var2(7),obj.rawData.Var3(7), ...
                               obj.rawData.Var4(7),obj.rawData.Var5(7), ...
                               obj.rawData.Var6(7),obj.rawData.Var7(7), ...
                               obj.rawData.Var8(7)];
            obj.SensorSerial = [obj.rawData.Var2(11),obj.rawData.Var3(11), ...
                                obj.rawData.Var4(11),obj.rawData.Var5(11), ...
                                obj.rawData.Var6(11),obj.rawData.Var7(11), ...
                                obj.rawData.Var8(11)];
            obj.engineeringUnit = [obj.rawData.Var2(14),obj.rawData.Var3(14), ...
                                   obj.rawData.Var4(14),obj.rawData.Var5(14), ...
                                   obj.rawData.Var6(14),obj.rawData.Var7(14), ...
                                   obj.rawData.Var8(14)];
            obj.sampleFreq = str2double(obj.rawData.Var2(5));
            obj.sampleTime = 1/obj.sampleFreq;  % Seconds per count

            % Get start and end of data
            obj.initVal = 24;   % This is where the data starts (usually 24 for the DTS)
            obj.endVal = size(obj.rawData, 1);

            
            % Get variable headers for table
            varHeads = obj.rawData.Properties.VariableNames;
            
            % Put data in array
            %obj.data = str2double(obj.rawData{24:obj.endVal, varHeads});
            obj.data = str2double(obj.rawData{24:obj.endVal,...
                {'Var2','Var3','Var4','Var5','Var6','Var7','Var8','Var1'}}); % here until I sort out the auto headers
            
            % Get numbers of rows and columns
            [obj.rows, obj.columns] = size(obj.data);
            
            % Set initial limits (These is no data outside these values)
            obj.cutStart = 0;
            obj.cutEnd =  obj.rows;
            
            % Set time column
            obj.time = obj.columns;
            
        end % Constructor obj = DataAnalysis(fileName)
        
        %% Create plots from file data
        function createPlots(obj)
            
            disp('Creating X v time plots...'); % Tell user whats going on
            %fprintf('Start limit: %d     End limit: %d\n', obj.cutStart, obj.cutEnd);
            
            % Function variables
            global ares fres cres mres;
            
            % Runtime variable control - set by user
            showPlots = obj.runTimeVals(1);     % Turns off graphs 
            smoothGraphs = obj.runTimeVals(2);  % Determine if smoothing is on
            smoothChoice = obj.runTimeVals(3);  % Set smoothing option
            smoothSpan = obj.runTimeVals(4);    % Determines smoothness
            timeConversion = obj.runTimeVals(5);% Conversion factor for time scale
            
            % Graph Control
            picSavepath = obj.plotVars(1);      % file path for save file
            picFileExt = obj.plotVars(2);       % file extension for save file
            picResolution = obj.plotVars(3);    % Plot resolution
            picFormatType = obj.plotVars(4);    % save format e.g. jpeg, png
            printGraphs = obj.plotVars(5);      % alows turning off plots
            plotControls = obj.plotSizing;      % Plots width height and screen position
            
            % If saving plots, create folder
            if ((exist([char(picSavepath),'\',char(obj.TestId)], 'dir') == 7))
               disp(['Saving to file ', char(picSavepath),'\',char(obj.TestId)]);
               %printGraphs = 'false';
               printGraphs = 'true';
            else
                if mkdir([char(picSavepath), '\',char(obj.TestId)])
                    disp(['Directory ',char(picSavepath), '\',char(obj.TestId),' created sucessfully']);
                    picSavepath = [char(picSavepath),'\', char(obj.TestId)];
                else
                    disp('Directory failed to create');
                end
            end
            
            % Get the type of graph smoothing
            switch smoothChoice
                case 0
                    smoothMethod = 'moving';
                case 1
                    smoothMethod = 'lowess';   
                case 2
                    smoothMethod = 'loess';
                case 3
                    smoothMethod = 'sgolay';
                case 4
                    smoothMethod = 'rlowess';
                case 5
                    smoothMethod = 'rloess';
                otherwise
                    smoothMethod = 'none';
            end

            % Allow plots to be turned on or off - mostly for debugging
            if showPlots == 1

                % Get axis resultant data
                if obj.headGx > 0 || obj.headGy > 0 || obj.headGz > 0
                    obj.resultantHeadAcel
                end
                if obj.neckFx > 0 || obj.neckFy > 0 || obj.neckFz > 0
                    obj.resultantNeckForce
                end
                if obj.neckMx > 0 || obj.neckMy > 0 || obj.neckMz > 0
                    obj.resultantNeckMoment
                end
                
                % time data - same for all graphs
                timeData = obj.data(obj.cutStart:obj.cutEnd, obj.time);
                
                % Creating the smoothed graphs
                if smoothGraphs == 1
                    fprintf('Smoothing method is %s\n', smoothMethod);
                    
                    % Head Acceleration
                    if obj.headGx > 0 || obj.headGy > 0 || obj.headGz > 0
                        
                        % Tell user whats going on
                        disp('Creating Head Acceleration plots');
                        
                        % Set Figure title
                        HeadAccelFig = figure ('Name', 'Head Acceleration',...
                                               'NumberTitle','off', 'pos', ...
                                                plotControls, 'visible','off');
                        
                        % Headx
                        if obj.headGx > 0
                            subplot(2,2,1)
                            HeadXY = smooth(timeData,obj.data(obj.cutStart:...
                                            obj.cutEnd,obj.headGx),smoothSpan,...
                                            smoothMethod);
                            plot(timeData*timeConversion,HeadXY)
                            xlabel('Time [ms]')
                            ylabel('Head x-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Heady
                        if obj.headGy > 0
                            subplot(2,2,2)
                            HeadYY = smooth(timeData,obj.data(obj.cutStart:...
                                            obj.cutEnd,obj.headGy),smoothSpan,...
                                            smoothMethod);
                            plot(timeData*timeConversion,HeadYY);
                            xlabel('Time [ms]')
                            ylabel('Head y-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Headz
                        if obj.headGz > 0
                            subplot(2,2,3)
                            HeadZY = smooth(timeData,obj.data(obj.cutStart:...
                                            obj.cutEnd,obj.headGz),smoothSpan,...
                                            smoothMethod);
                            plot(timeData*timeConversion,HeadZY);
                            xlabel('Time [ms]')
                            ylabel('Head z-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Head resultant
                        if obj.headGx > 0 || obj.headGy > 0 || obj.headGz > 0
                            subplot(2,2,4)
                            ChestResY = smooth(obj.data(obj.cutStart:...
                                               obj.cutEnd, obj.time),...
                                               ares(obj.cutStart:obj.cutEnd),...
                                               smoothSpan,smoothMethod);
                            plot( timeData*timeConversion, ChestResY )
                            xlabel('Time [ms]')
                            ylabel('Resultant Head Acceleration [g]')
                            grid on
                            zoom on
                        end
                        
                        % print pictures of the graphs - Head Acceleration
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId),...
                                         '\' , char(obj.TestId) ,'_HeadAccelSmooth',...
                                         char(picFileExt)];
                            fprintf('Saving graph for head acceleration at: %s\n', ...
                                     printPath);
                            print(HeadAccelFig, printPath, char(picResolution), ...
                                  char(picFormatType));
                        else
                            disp('head Accel was not saved');
                        end
                    end
                    
                    % Chest Acceleration
                    if obj.chestGx > 0 || obj.chestGy > 0 || obj.chestGz > 0
                        % Tell user whats going on
                        disp('Creating Chest Acceleration plots');
                        
                        % Set Figure title
                        ChestAccelFig = figure ('Name', 'Head Acceleration',...
                                                'NumberTitle','off', 'pos', ...
                                                plotControls, 'visible','off');
                        
                        % Chestx
                        if obj.chestGx > 0
                            subplot(2,2,1)
                            ChestXY = smooth(timeData,obj.data(obj.cutStart:...
                                             obj.cutEnd,obj.chestGx),smoothSpan,...
                                             smoothMethod);
                            plot(timeData*timeConversion,ChestXY)
                            xlabel('Time [ms]')
                            ylabel('Chest x-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Chesty
                        if obj.chestGy > 0
                            subplot(2,2,2)
                            ChestYY = smooth(timeData,obj.data(obj.cutStart:...
                                             obj.cutEnd,obj.chestGy),smoothSpan,...
                                             smoothMethod);
                            plot(timeData*timeConversion,ChestYY);
                            xlabel('Time [ms]')
                            ylabel('Chest y-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Chestz
                        if obj.chestGz > 0
                            subplot(2,2,3)
                            ChestZY = smooth(timeData,obj.data(obj.cutStart:...
                                             obj.cutEnd,obj.chestGz),smoothSpan,...
                                             smoothMethod);
                            plot(timeData*timeConversion,ChestZY);
                            xlabel('Time [ms]')
                            ylabel('Chest z-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Chest resultant
                        if obj.chestGx > 0 || obj.chestGy > 0 || obj.chestGz > 0
                            subplot(2,2,4)
                            ChestResY = smooth(obj.data(obj.cutStart:obj.cutEnd,...
                                               obj.time),cres(obj.cutStart:...
                                               obj.cutEnd),smoothSpan,smoothMethod);
                            plot( timeData*timeConversion, ChestResY )
                            xlabel('Time [ms]')
                            ylabel('Resultant Chest Acceleration [g]')
                            grid on
                            zoom on
                        end
                        
                        % print pictures of the graphs - Chest Acceleration
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_ChestAccelSmooth', char(picFileExt)];
                            fprintf('Saving graph for chest acceleration at: %s\n', printPath);
                            print(ChestAccelFig, printPath, char(picResolution), char(picFormatType));
                        end
                    end
                    
                    % Neck Forces
                    if obj.neckFx > 0 || obj.neckFy > 0 || obj.neckFz > 0
                        % Tell user whats going on
                        disp('Creating Neck Force plots');
                        
                        % Set Figure title
                        NeckForceFig = figure ('Name', 'Neck Forces',...
                                               'NumberTitle','off', 'pos',...
                                                plotControls, 'visible','off');
                        
                        % NeckFx
                        if obj.neckFx > 0
                            subplot(2,2,1)
                            NeckFxY = smooth(obj.data(obj.cutStart:obj.cutEnd,...
                                             obj.time),obj.data(obj.cutStart:...
                                             obj.cutEnd,obj.neckFx),smoothSpan,...
                                             smoothMethod);
                            plot(timeData*timeConversion,NeckFxY);
                            xlabel('Time [ms]')
                            ylabel('Upper neck x-force [kN]')
                            grid on
                            zoom on
                        end

                        % NeckFy
                        if obj.neckFy > 0                   
                            subplot(2,2,2)
                            NeckFyY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.neckFy),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,NeckFyY);
                            xlabel('Time [ms]')
                            ylabel('Upper neck y-force [kN]')
                            grid on
                            zoom on
                        end

                        % NeckFz
                        if obj.neckFz > 0
                            subplot(2,2,3)
                            NeckFzY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.neckFz),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,NeckFzY);
                            xlabel('Time [ms]')
                            ylabel('Upper Neck z-force [kN]')
                            grid on
                            zoom on
                        end

                        % NeckFRes
                        if obj.neckFx > 0 || obj.neckFy > 0 || obj.neckFz > 0
                            subplot(2,2,4)
                            neckFResY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),fres(obj.cutStart:obj.cutEnd),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,neckFResY);
                            xlabel('Time [ms]')
                            ylabel('Upper Neck Resultant Force [N]')
                            grid on
                            zoom on
                        end
                        
                        % print pictures of the graphs - Neck Acceleration
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_NeckforceSmooth', char(picFileExt)];
                            fprintf('Saving graph for neck forces at: %s\n', printPath);
                            print(NeckForceFig, printPath, char(picResolution), char(picFormatType));
                        end
                    end
                    
                    % Neck Moment
                    if obj.neckMx > 0 || obj.neckMy > 0 || obj.neckMz > 0
                        % Tell user whats going on
                        disp('Creating Neck Moment plots');
                        
                        % Set Figure title
                        NeckMomentFig = figure ('Name', 'Neck Moment', 'NumberTitle','off', 'pos', plotControls, 'visible','off');
                        
                        % NeckMx
                        if obj.neckMx > 0
                            subplot(2,2,1)
                            plot(timeData*timeConversion,obj.data(cutFront:cutRear,obj.neckMx));
                            xlabel('Time [ms]')
                            ylabel('Upper neck x-moment [Nm]')
                            grid on
                            zoom on
                        end

                        % NeckMy
                        if obj.neckMy > 0
                            subplot(2,2,2)
                            plot(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.neckMy));
                            xlabel('Time [ms]')
                            ylabel('Upper neck y-moment [Nm]')
                            grid on
                            zoom on
                        end

                        % NeckMz
                        if obj.neckMz > 0
                            subplot(2,2,3)
                            plot(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.neckMz));
                            xlabel('Time [ms]')
                            ylabel('Upper neck z-moment [Nm]')
                            grid on
                            zoom on
                        end

                        % NeckMRes
                        if obj.neckMx > 0 || obj.neckMy > 0 || obj.neckMz > 0
                            subplot(2,2,4)
                            neckMResY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),mres(obj.cutStart:obj.cutEnd),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,neckMResY);
                            xlabel('Time [ms]')
                            ylabel('Upper neck Resultant [Nm]')
                            grid on
                            zoom on
                        end
                        
                        % print pictures of the graphs - Neck moment
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_NeckmomentSmooth', char(picFileExt)];
                            fprintf('Saving graph for neck moment at: %s\n', printPath);
                            print(NeckMomentFig, printPath, char(picResolution), char(picFormatType));
                        end
                    end
                    
                    % Iliac LC
                    if obj.iliacLFx > 0 || obj.iliacLMy > 0 || obj.iliacRFx > 0 || obj.iliacRMy > 0
                        % Tell user whats going on
                        disp('Creating Iliac LC plots');
                        
                        % Set Figure title
                        IliacLCFig = figure ('Name', 'Iliac LC', 'NumberTitle','off', 'pos', plotControls, 'visible','off');
                        
                        % Iliac Left Fx
                        if obj.iliacLFx > 0
                            subplot(2,2,1)
                            IliacLFxY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.iliacLFx),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,IliacLFxY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Left Fx [kN]')
                            grid on
                            zoom on
                        end
                        
                        % Iliac Left Fx
                        if obj.iliacLMy > 0
                            subplot(2,2,2)
                            IliacLMyY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.iliacLMy),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,IliacLMyY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Left My [Nm]')
                            grid on
                            zoom on
                        end
                        
                        % Iliac Left Fx
                        if obj.iliacRFx > 0
                            subplot(2,2,3)
                            IliacRFxY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.iliacRFx),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,IliacRFxY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Right Fx [kN]')
                            grid on
                            zoom on
                        end
                        
                        % Iliac Left Fx
                        if obj.iliacRMy > 0
                            subplot(2,2,4)
                            IliacRMyY = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),obj.data(obj.cutStart:obj.cutEnd,obj.iliacRMy),smoothSpan,smoothMethod);
                            plot(timeData*timeConversion,IliacRMyY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Right My [Nm]')
                            grid on
                            zoom on
                        end
                        
                        % print pictures of the graphs - Iliac LC
                        if strcmpi(string(printGraphs),'true')    
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_IliacLCSmooth', char(picFileExt)];
                            fprintf('Saving graph for Iliac LC at: %s\n', printPath);
                            print(IliacLCFig, printPath, char(picResolution), char(picFormatType));
                        end
                    end
                else
                    % Non-smoothed graphs
                    % Head Acceleration
                    if obj.headGx > 0 || obj.headGy > 0 || obj.headGz > 0
                        % Tell user whats going on
                        disp('Creating Head Acceleration plots');
                        
                        % Set Figure title
                        HeadAccelFig = figure ('Name', 'Head Acceleration', 'NumberTitle','off', 'pos', plotControls, 'visible','off');

                        % HeadGx
                        if obj.headGx > 0
                            subplot(2,2,1)
                            HeadXY = obj.data(obj.cutStart:obj.cutEnd,obj.headGx);
                            plot(timeData*timeConversion,HeadXY)
                            xlabel('Time [ms]')
                            ylabel('Head x-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Heady
                        if obj.headGy > 0
                            subplot(2,2,2)
                            HeadYY = obj.data(obj.cutStart:obj.cutEnd,obj.headGy);
                            plot(timeData*timeConversion,HeadYY);
                            xlabel('Time [ms]')
                            ylabel('Head y-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Headz
                        if obj.headGz > 0
                            subplot(2,2,3)
                            HeadZY = obj.data(obj.cutStart:obj.cutEnd,obj.headGz);
                            plot(timeData*timeConversion,HeadZY);
                            xlabel('Time [ms]')
                            ylabel('Head z-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Head resultant
                        if obj.headGx > 0 || obj.headGy > 0 || obj.headGz > 0
                            subplot(2,2,4)
                            ChestResY = ares(obj.cutStart:obj.cutEnd);
                            plot(timeData*timeConversion, ChestResY )
                            xlabel('Time [ms]')
                            ylabel('Resultant Head Acceleration [g]')
                            grid on
                            zoom on
                        end

                        % print pictures of the graphs - Head Acceleration
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_HeadAccel', char(picFileExt)];
                            fprintf('Saving graph for head acceleration at: %s\n', printPath);
                            print(HeadAccelFig, printPath, char(picResolution), char(picFormatType));
                        else
                            disp('head Accel was not saved');
                        end
                    end
                        
                    % Chest Acceleration
                    if obj.chestGx > 0 || obj.chestGy > 0 || obj.chestGz > 0
                        % Tell user whats going on
                        disp('Creating Chest Acceleration plots');
                        
                        % Set Figure title                        
                        ChestAccelFig = figure ('Name', 'Head Acceleration', 'NumberTitle','off', 'pos', plotControls, 'visible','off');
                        
                        % Chestx
                        if obj.chestGx > 0
                            subplot(2,2,1)
                            ChestXY = obj.data(obj.cutStart:obj.cutEnd,obj.chestGx);
                            plot(timeData*timeConversion,ChestXY)
                            xlabel('Time [ms]')
                            ylabel('Chest x-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Chesty
                        if obj.chestGy > 0
                            subplot(2,2,2)
                            ChestYY = obj.data(obj.cutStart:obj.cutEnd,obj.chestGy);
                            plot(timeData*timeConversion,ChestYY);
                            xlabel('Time [ms]')
                            ylabel('Chest y-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Chestz
                        if obj.chestGz > 0
                            subplot(2,2,3)
                            ChestZY = obj.data(obj.cutStart:obj.cutEnd,obj.chestGz);
                            plot(timeData*timeConversion,ChestZY);
                            xlabel('Time [ms]')
                            ylabel('Chest z-acceleration [G]')
                            grid on
                            zoom on
                        end

                        % Chest resultant
                        if obj.chestGx > 0 || obj.chestGy > 0 || obj.chestGz > 0
                            subplot(2,2,4)
                            ChestResY = cres(obj.cutStart:obj.cutEnd);
                            plot( timeData*timeConversion, ChestResY )
                            xlabel('Time [ms]')
                            ylabel('Resultant Chest Acceleration [g]')
                            grid on
                            zoom on
                        end

                        % print pictures of the graphs - Chest Acceleration
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_ChestAccel', char(picFileExt)];
                            fprintf('Saving graph for chest acceleration at: %s\n', printPath);
                            print(ChestAccelFig, printPath, char(picResolution), char(picFormatType));
                        else
                            disp('Chest Accel was not saved');
                        end
                    end

                    % Neck Forces
                    if obj.neckFx > 0 || obj.neckFy > 0 || obj.neckFz > 0
                        % Tell user whats going on
                        disp('Creating Neck Force plots');
                        
                        % Set Figure title
                        NeckForceFig = figure ('Name', 'Neck Forces', 'NumberTitle','off', 'pos', plotControls, 'visible','off');
                        
                        % NeckFx
                        if obj.neckFx > 0
                            subplot(2,2,1)
                            plot(timeData*timeConversion,obj.data(obj.cutStart:obj.cutEnd,obj.neckFx));
                            xlabel('Time [ms]')
                            ylabel('Upper neck x-force [N]')
                            grid on
                            zoom on
                        end

                        % NeckFy
                        if obj.neckFy > 0
                            subplot(2,2,2)
                            plot(timeData*timeConversion,obj.data(obj.cutStart:obj.cutEnd,obj.neckFy));
                            xlabel('Time [ms]')
                            ylabel('Upper neck y-force [N]')
                            grid on
                            zoom on
                        end

                        % NeckFz
                        if obj.neckFz > 0
                            subplot(2,2,3)
                            plot(timeData*timeConversion,obj.data(obj.cutStart:obj.cutEnd,obj.neckFz));
                            xlabel('Time [ms]')
                            ylabel('Upper neck z-force [N]')
                            grid on
                            zoom on
                        end

                        % NeckFRes
                        if obj.neckFx > 0 || obj.neckFy > 0 || obj.neckFz > 0
                            subplot(2,2,4)
                            plot(timeData*timeConversion,ares(obj.cutStart:obj.cutEnd));
                            xlabel('Time [ms]')
                            ylabel('Upper Neck Resultant Force [N]')
                            grid on
                            zoom on
                        end

                        % print pictures of the graphs - Neck Force
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_Neckforce', char(picFileExt)];
                            fprintf('Saving graph for neck forces at: %s\n', printPath);
                            print(NeckForceFig, printPath, char(picResolution), char(picFormatType));
                        else
                            disp('Neck force was not saved');
                        end
                    end
                    
                    % Neck Moment
                    if obj.neckMx > 0 || obj.neckMy > 0 || obj.neckMz > 0
                        % Tell user whats going on
                        disp('Creating Neck Moment plots');
                        
                        % Set Figure title
                        NeckMomentFig = figure ('Name', 'Neck Moment', 'NumberTitle','off', 'pos', plotControls, 'visible','off');
                        
                        % NeckMx
                        if obj.neckMx > 0
                            subplot(2,2,1)
                            plot(timeData*timeConversion,obj.data(cutFront:cutRear,obj.neckMx));
                            xlabel('Time [ms]')
                            ylabel('Upper neck x-moment [Nm]')
                            grid on
                            zoom on
                        end

                        % NeckMy
                        if obj.neckMy > 0
                            subplot(2,2,2)
                            plot(timeData*timeConversion,obj.data(obj.cutStart:obj.cutEnd,obj.neckMy));
                            xlabel('Time [ms]')
                            ylabel('Upper neck y-moment [Nm]')
                            grid on
                            zoom on
                        end

                        % NeckMz
                        if obj.neckMz > 0
                            subplot(2,2,3)
                            plot(timeData*timeConversion,obj.data(obj.cutStart:obj.cutEnd,obj.neckMz));
                            xlabel('Time [ms]')
                            ylabel('Upper neck z-moment [Nm]')
                            grid on
                            zoom on
                        end

                        % NeckMRes
                        if obj.neckMx > 0 || obj.neckMy > 0 || obj.neckMz > 0
                            subplot(2,2,4)
                            plot(timeData*timeConversion,mres(obj.cutStart:obj.cutEnd));
                            xlabel('Time [ms]')
                            ylabel('Upper neck Resultant [Nm]')
                            grid on
                            zoom on
                        end

                        % print pictures of the graphs - Neck moment
                        if strcmpi(string(printGraphs),'true')
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_Neckmoment', char(picFileExt)];
                            fprintf('Saving graph for neck moment at: %s\n', printPath);
                            print(NeckMomentFig, printPath, char(picResolution), char(picFormatType));
                        else
                            disp('Neck moment was not saved');
                        end

                    end
                    
                    % Iliac LC
                    if obj.iliacLFx > 0 || obj.iliacLMy > 0 || obj.iliacRFx > 0 || obj.iliacRMy > 0
                        % Tell user whats going on
                        disp('Creating Iliac LC plots');
                        
                        % Set Figure title
                        IliacLCFig = figure ('Name', 'Iliac LC', 'NumberTitle','off', 'pos', plotControls, 'visible','off');
                        
                        % Iliac Left Fx
                        if obj.iliacLFx > 0
                            subplot(2,2,1)
                            IliacLFxY = obj.data(obj.cutStart:obj.cutEnd,obj.iliacLFx);
                            plot(timeData*timeConversion,IliacLFxY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Left Fx [kN]')
                            grid on
                            zoom on
                        end
                        
                        % Iliac Left Fx
                        if obj.iliacLMy > 0
                            subplot(2,2,2)
                            IliacLMyY = obj.data(obj.cutStart:obj.cutEnd,obj.iliacLMy);
                            plot(timeData*timeConversion,IliacLMyY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Left My [Nm]')
                            grid on
                            zoom on
                        end
                        
                        % Iliac Left Fx
                        if obj.iliacRFx > 0
                            subplot(2,2,3)
                            IliacRFxY = obj.data(obj.cutStart:obj.cutEnd,obj.iliacRFx);
                            plot(timeData*timeConversion,IliacRFxY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Right Fx [kN]')
                            grid on
                            zoom on
                        end
                        
                        % Iliac Left Fx
                        if obj.iliacRMy > 0
                            subplot(2,2,4)
                            IliacRMyY = obj.data(obj.cutStart:obj.cutEnd,obj.iliacRMy);
                            plot(timeData*timeConversion,IliacRMyY);
                            xlabel('Time [ms]')
                            ylabel('Iliac Right My [Nm]')
                            grid on
                            zoom on
                        end
                        
                        % print pictures of the graphs - Iliac LC
                        if strcmpi(string(printGraphs),'true')    
                            printPath = [char(picSavepath),'\' , char(obj.TestId) ,'\' , char(obj.TestId) ,'_IliacLC', char(picFileExt)];
                            fprintf('Saving graph for Iliac LC at: %s\n', printPath);
                            print(IliacLCFig, printPath, char(picResolution), char(picFormatType));
                        end
                    end
                end
            else
                disp('ShowPlots is off')
            end 
            disp('Finished creating plots')
        end % obj = creatPlots
        
        %% Resultant Head Acceleration
        function resultantHeadAcel(obj)
            
            % variables
            global ares;
            
            % Tell user whats going on
            disp('Calculating resultant head acceleration...'); %Debug
            
            % Get data for each axis
            ax = obj.data(:,obj.headGx);
            ay = obj.data(:,obj.headGy);
            az = obj.data(:,obj.headGz);
            
            % Get resultant
            ares = sqrt((ax.^2) + (ay.^2) + (az.^2));
            
            % Turn off resultant plot here - only for debug
            showPlots = 0;
            
            % Show plots if necessary
            if showPlots == 1
                figure('Name', 'Resultant Head Acceleration [g]', 'NumberTitle','off')
                
                % Allow smoothing of plot
                smoothed = 0;
                x = obj.data(obj.cutStart:obj.cutEnd, obj.time)*1000;
                if smoothed == 1
                    fprintf('Smoothing data\n');
                    %smoothmethod = 'loess';    % nick(get these to use user data please)
                    smoothmethod = 'rloess';
                    span = 0.03;    % Determines smoothness
                    smoothedData = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),ares(obj.cutStart:obj.cutEnd),span,smoothmethod);
                    
                    y = smoothedData;
                    plot( x, y )
                    xlabel('Time [ms]')
                    ylabel('Resultant Head Acceleration [g]')
                    grid on
                    zoom on
                else
                    plot( x, ares(obj.cutStart:obj.cutEnd, 1))
                    xlabel('Time [ms]')
                    ylabel('Resultant Head Acceleration [g]') 
                    grid on
                    zoom on
                end
            else
                %disp ('Resultant head accel showPlots is off')
            end
            fprintf('Calculate Resultant Head Accleration complete\n');
        end % obj = resultantHeadAcel(data)
        
        %% Resultant Chest Acceleration
        function resultantChestAcel(obj)
            
            disp('Calculating resultant chest acceleration...');
            ax = obj.data(:,obj.chestGx);
            ay = obj.data(:,obj.chestGy);
            az = obj.data(:,obj.chestGz);
            global cres;
            cres = sqrt((ax.^2) + (ay.^2) + (az.^2));
            %obj.headAccelRes = ares;
            
            showPlots = 0;
            
            if showPlots == 1
                figure('Name', 'Resultant Chest Acceleration [g]', 'NumberTitle','off')
                
                % Allow smoothing of plot
                smoothed = 1;
                x = obj.data(obj.cutStart:obj.cutEnd, obj.time)*1000;
                if smoothed == 1
                    fprintf('Smoothing data\n');
                    %smoothmethod = 'loess';
                    smoothmethod = 'rloess';
                    span = 0.03;    % Determines smoothness
                    smoothedData = smooth(obj.data(obj.cutStart:obj.cutEnd, obj.time),cres(obj.cutStart:obj.cutEnd),span,smoothmethod);
                    
                    y = smoothedData;
                    plot( x, y )
                    xlabel('Time [ms]')
                    ylabel('Resultant Chest Acceleration [g]')
                    grid on
                    zoom on
                else
                    plot( x, cres(obj.cutStart:obj.cutEnd, 1))
                    xlabel('Time [ms]')
                    ylabel('Resultant Chest Acceleration [g]') 
                    grid on
                    zoom on
                end
            else
                %disp ('Resultant chest accel showPlots is off')
            end
            fprintf('Calculate Resultant Chest Accleration complete\n');
        end % obj = resultantChestAcel(data)
        
        %% Calculates the resultant neck forces
        function resultantNeckForce(obj)
            %%% RESULTANT UPPER NECK FORCES [N]
            disp('Calculating resultant neck forces...'); %Debug
            fx = 0;
            fy = 0;
            fz = 0;
            % make sure there is a value for each
            if obj.neckFx > 0
                fx = obj.data(:,obj.neckFx);
            end
            if obj.neckFy > 0
                fy = obj.data(:,obj.neckFy);
            end
            if obj.neckFz > 0
                fz = obj.data(:,obj.neckFz);
            end
            
            global fres;
            fres = sqrt((fx.^2) + (fy.^2) + (fz.^2));
            
            showPlots = 0;
            
            if showPlots == 1
            
                figure('Name', 'Resultant upper neck force [N]', 'NumberTitle','off')
                plot( obj.data(obj.cutStart:obj.cutEnd, obj.time), fres(obj.cutStart:obj.cutEnd))
                xlabel('Time [ms]')
                ylabel('Resultant upper neck force [N]')
                grid on
                zoom on
            else
                %disp('Resultant neck forces showPlots off')
            end
            fprintf('Calculate Resultant Neck Forces complete\n');
        end % obj = resultantNeckForce(data)
        
        %% Calculates the resultant neck moments
        function resultantNeckMoment(obj)
            %%% RESULTANT UPPER NECK MOMENTS [Nm]
            disp('Calculating resultant neck moments...'); %Debug
            
            mx = 0;
            my = 0;
            mz = 0;
            % make sure there is a value for each
            if obj.neckMx > 0
                mx = obj.data(:,obj.neckMx);
            end
            if obj.neckMy > 0
                my = obj.data(:,obj.neckMy);
            end
            if obj.neckMz > 0
                mz = obj.data(:,obj.neckMz);
            end
            
            global mres;
            mres = sqrt((mx.^2) + (my.^2) + (mz.^2));
            
            showPlots = 0;
            
            if showPlots == 1
            
                figure('Name', 'Resultant upper neck moment [Nm]', 'NumberTitle','off')
                plot( obj.data(obj.cutStart:obj.cutEnd, obj.time), mres(obj.cutStart:obj.cutEnd))
                xlabel('Time [ms]')
                ylabel('Resultant upper neck moment [Nm]')
                grid on
                zoom on
            else
              %disp('Resultant neck moments showPlots off')
            end
            
            disp('Finished calculating resultant neck moments');
            
        end
 
        %% Calculates the head injury criteria
        function calcHIC(obj)
            %% HEAD INJURY CRITERIA (HIC)
            
            disp('Calculating HIC...'); %Debug
            
            % HIC is calculated as the maximum value of an equation including 
            % t_1, t_2 and resultant acceleration, where t_1 and t_2 are 
            % separated by no more than 36 msec. 
            
            % HIC formula
            %                   /t2
            % HIC = {[1/(t2-t1) |  a(t)dt]^2.5 * (t2-t1))max
            %                   /t1
            
            global ares;
            
            % get size of head accleration vector
            headAcelVec = size(ares);
            
            % Calculate resultant head acceleration if not done already
            if headAcelVec == 0
                obj.resultantHeadAcel;
                headAcelVec = size(ares);
            end

            % initialise HIC value
            HIC15 = 0; 

            % HIC15 calculation 
            % iteration values
             startHIC15 = obj.cutStart;
             endHIC15 = startHIC15 + (1.5*str2double(obj.sampleRate(2)));    %Auto calc 15ms from start point
             stepHIC15 = 1000; 
             %tick1 = 0;
             %tick2 = 0;
             
             disp('This may take a while. Go have coffee and a biscuit =-)');
             disp('Calculating HIC15');
             for outerIt15 = startHIC15:stepHIC15:endHIC15   % 150 = 15 msec
                % lastIt = Last iteration to be made (15ms before end of vector) 
                lastIt = headAcelVec - (outerIt15+1);
               
               for innerIt15 = 1:lastIt
                  t_1 = innerIt15;
                  t_2 = innerIt15+outerIt15;
                  currentset = ares(t_1:t_2);   % Vector of acceleration points spanning jvalue (*0.1 msec)
                  area = cumsum(obj.sampleTime*currentset); % area (or integral) calculated in seconds 
                  [ x, y ] = size(area);
                  integral = area(x,y);
                  % HIC equation calculated for timespan jvalue (seconds) %(*0.1 msec)
                  currentvalue = ((integral/(obj.sampleTime*(t_2-t_1)))^2.5)*(t_2-t_1)*obj.sampleTime; 
                  %tick1 = tick1+1;
                  %disp(currentvalue)
                  if currentvalue > HIC15 
                     HIC15 = currentvalue;
                     %ivalue15 = innerIt15;
                     %jvalue15 = (t_2 - t_1);
                  end
                  %tick2 = tick2+1;
               end
               
               % Progress bar
               if mod(outerIt15,stepHIC15) == 0
                    fprintf('|');
               end
             end           
             fprintf('\n');
            %tick1
            %tick2
            
            % HIC36 calculation
            HIC36 = HIC15;
            %ivalue36 = ivalue15;
            %jvalue36 = jvalue15;
            % iteration values
             %startHIC36 = 10025; % need to get this from file
             endHIC36 = startHIC15 + (3.6*str2double(obj.sampleRate(2)));    %Auto calc 36ms from start point
             stepHIC36 = stepHIC15; 
            
            disp('Calculating HIC36');
            for outerIt36 = endHIC15:stepHIC36:endHIC36   % 360 = 36 msec
               lastIt = headAcelVec - (outerIt36+1);   % k = Last iteration to be made (36ms before end of vector) 
               for innerIt36 = 1:lastIt
                  t_1 = innerIt36;
                  t_2 = innerIt36+outerIt36;
                  currentset = ares(t_1:t_2);   % Vector of acceleration points spanning jvalue (*0.1 msec)
                  area = cumsum(obj.sampleTime*currentset); % area (or integral) calculated in seconds 
                  [ x, y ] = size(area);
                  integral = area(x,y);            
                  currentvalue = ((integral/(obj.sampleTime*(t_2-t_1)))^2.5)*(t_2-t_1)*obj.sampleTime;  % HIC equation calculated for timespan jvalue (seconds) %(*0.1 msec)
                  if currentvalue > HIC36 
                     HIC36 = currentvalue; 
                     %ivalue36 = innerIt36;
                     %jvalue36 = (t_2 - t_1);
                  end
               end
              % Loading bar
               if mod(outerIt36,stepHIC36) == 0
                    fprintf('|');
               end
            end
            fprintf('\n');
%             for j = 150:10:360   % 360 = 36 msec
%                lastIt = headAcelVec - (j+1);   % k = Last iteration to be made (36ms before end of vector) 
%                for innerIt = 1:lastIt
%                   t_1 = innerIt;
%                   t_2 = innerIt+j;
%                   currentset = ares(t_1:t_2);   % Vector of acceleration points spanning jvalue (*0.1 msec)
%                   area = cumsum(obj.sampleTime*currentset); % area (or integral) calculated in seconds 
%                   [ x, y ] = size(area);
%                   integral = area(x,y);            
%                   currentvalue = ((integral/(obj.sampleTime*(t_2-t_1)))^2.5)*(t_2-t_1)*obj.sampleTime;  % HIC equation calculated for timespan jvalue (seconds) %(*0.1 msec)
%                   if currentvalue > HIC36 
%                      HIC36 = currentvalue; 
%                      %ivalue36 = i;
%                      %jvalue36 = (t_2 - t_1);
%                   end
%               end
%             end
            

            fprintf('HIC15 is: %.3f\n', HIC15);
            obj.HICVal15 = HIC15;
            %t1_15 = ivalue15 * 0.1; 
            %t2_15 = t1_15 + jvalue15 *0.1;
            %disp('HIC36 is: ');
            fprintf('HIC36 is: %.3f\n', HIC36);
            obj.HICVal36 = HIC36;
            %t1_36 = ivalue36 * 0.1; 
            %t2_36 = t1_36 + jvalue36 *0.1; 
            disp('Finished calculating HIC')
        end % obj = calcHIC(data)
        
        %% Calculates the neck injury criteria
        function calcNeckInjuryCriteria(obj)
                %%% NECK INJURY CRITERIA (Nij)
                disp('Calculating neck injury criteria...'); %Debug
                % For current FMVSS no.208 standards: 
                % Critical intercept values used for normalisation.
                % Fzt_crit - neck tension force, positive Fz 
                % Fzc_crit - neck compression force, negative Fz
                % Myf_crit - neck flexion moment, positive My
                % Mye_crit - neck extension moment, negative My
                %crit = [Fzt  Fzc  Myf Mye];
                %crit = [2200 2200  85 25]; %CRABI child dummy 12 months
                %crit = [2500 2500 100 30]; %Hybrid III child dummy 3 years
                crit = [2900 2900 125 40]; %Hybrid III child dummy 6 years
                Fzt_crit = crit(1);   
                Fzc_crit = crit(2);   
                Myf_crit = crit(3);     
                Mye_crit = crit(4);     

                % Defining force and moment vectors from data file (from beginning to end of pulse) 
                % Fz into tension(t) and compression(c) and My into extension(e) and flexion(f)
                Fz = obj.data(obj.cutStart:obj.cutEnd,7);   
                My = obj.data(obj.cutStart:obj.cutEnd,8);     
                Fz_abs = abs(Fz);
                Fzt = (Fz + Fz_abs) / 2;
                Fzc = (Fz_abs - Fz) / 2;
                My_abs = abs(My);
                Myf = (My + My_abs) / 2;
                Mye = (My_abs - My) / 2;

                figure ('Name', 'Neck Injruy Criteria - Fz', 'NumberTitle','off')
                
                %obj.data(obj.cutFront:obj.cutRear, obj.time),obj.data(obj.cutFront:obj.cutRear,obj.neckMy
                
                plot ( obj.data(obj.cutStart:obj.cutEnd, obj.time), Fzt, obj.data(obj.cutStart:obj.cutEnd, obj.time), Fzc)
                xlabel('t [ms]')
                ylabel('Upper neck axial force [N]')
                legend('Fz Tension','Fz Compression')
                grid on
                zoom on

                figure ('Name', 'Neck Injruy Criteria - My - ', 'NumberTitle','off')
                plot ( obj.data(obj.cutStart:obj.cutEnd, obj.time), Myf, obj.data(obj.cutStart:obj.cutEnd, obj.time), Mye)
                xlabel('t [ms]')
                ylabel('Upper neck y-moment [Nm]')
                legend('My Flexion','My Extension')
                grid on
                zoom on

                % Calculate all 4 Nij curves and the maximum Nij curve for the whole sled pulse 
                NTF = Fzt/Fzt_crit + Myf/Myf_crit;  % Neck tension / flexion
                NTE = Fzt/Fzt_crit + Mye/Mye_crit;  % Neck tension / extension
                NCF = Fzc/Fzc_crit + Myf/Myf_crit;  % Neck compression / flexion
                NCE = Fzc/Fzc_crit + Mye/Mye_crit;  % Neck compression / extension 
                %Nij = [NTF NTE NCF NCE];
               
                %[Nmax,iNmax] = max( Nij(:),[], 1 );
                %disp ('    NTF       NTE       NCF       NCE')
                %Nijmax = Nmax
                %0.1*iNmax

                figure ('Name', 'Neck Injruy Criteria - Neck Tension', 'NumberTitle','off')
                plot( obj.data(obj.cutStart:obj.cutEnd, obj.time), NTF, ...
                      obj.data(obj.cutStart:obj.cutEnd, obj.time), NTE, ...
                      obj.data(obj.cutStart:obj.cutEnd, obj.time), NCF, ...
                      obj.data(obj.cutStart:obj.cutEnd, obj.time), NCE)
                xlabel('Time [ms]')
                ylabel('Nij')
                legend('NTF','NTE','NCF','NCE')
                grid on
                zoom on

                disp('Finished calcuating NIC');
        end
           
         %% Retreves the maximum values in the data
         function calcMaxVals(obj)
                %%% MAXIMUM VALUES
                disp('Calculating max values...'); %Debug
                
                % Variables
                maxVal = {}; % Cell array
                index = 1;
                global ares fres mres;
                
                % try using an index e.g 
                % 
                % maxVal{index,1} = 'Sled Max';
                % index = index + 1;
                
                % SledG max
                if obj.sledG > 0
                    [sledxmax,~] = max(obj.data(:,obj.sledG));
                    maxVal{index,1} = 'Sled Max (G)';
                    maxVal{index,2} = sledxmax;
                    index = index + 1;
                else
                    %sledxmax = 0;
                end
                
                % HeadGx max
                if obj.headGx > 0
                    [axmax,~] = max(obj.data(:,obj.headGx));
                    maxVal{index,1} = 'HeadGx (G)';
                    maxVal{index,2} = axmax;
                    index = index + 1;
                else
                    %axmax = 0;
                end
                
                % HeadGy max
                if obj.headGy > 0
                    [aymax,~] = max(obj.data(:,obj.headGy));
                    maxVal{index,1} = 'HeadGy (G)';
                    maxVal{index,2} = aymax;
                    index = index + 1;
                else
                    % = 0;
                end
                
                % HeadGz max
                if obj.headGz > 0
                    [azmax,~] = max(obj.data(:,obj.headGz));
                    maxVal{index,1} = 'HeadGz (G)';
                    maxVal{index,2} = azmax;
                    index = index + 1;
                else
                    %azmax = 0;
                end
                
                % Head Resultant max
                if ares > 0
                    [aresmax,~] = max(ares);
                    maxVal{index,1} = 'Head Accel Resultant (G)';
                    maxVal{index,2} = aresmax;
                    index = index + 1;
                else
                    %aresmax = 0;
                end
                  
                % NeckFx max
                if obj.headGz > 0
                    [fxmax,~] = max(obj.data(:,obj.neckFx));
                    maxVal{index,1} = 'NeckFx (kN)';
                    maxVal{index,2} = fxmax;
                    index = index + 1;
                else
                    %fxmax = 0;
                end
                
                % NeckFy max
                if obj.neckFy > 0
                    [fymax,~] = max(obj.data(:,obj.neckFy));
                    maxVal{index,1} = 'NeckFy (kN)';
                    maxVal{index,2} = fymax;
                    index = index + 1;
                else
                    %fymax = 0;
                end
                
                % NeckFz max
                if obj.neckFz > 0
                    [fzmax,~] = max(obj.data(:,obj.neckFz));
                    maxVal{index,1} = 'neckFz (kN)';
                    maxVal{index,2} = fzmax;
                    index = index + 1;
                else
                    %fzmax = 0;
                end
                
                % Neck Force Resultant max
                if fres > 0
                    [fresmax,~] = max(fres);
                    maxVal{index,1} = 'Neck Force Resultant (kN)';
                    maxVal{index,2} = fresmax;
                    index = index + 1;
                else
                    %fresmax = 0;
                end
                
                % NeckMx max
                if obj.neckMx > 0
                    mxmax = max(obj.data(:,obj.neckMx));
                    maxVal{index,1} = 'NeckMx (Nm)';
                    maxVal{index,2} = mxmax;
                    index = index + 1;
                else
                    %mxmax = 0;
                end
                
                % NeckMy max
                if obj.neckMy > 0
                    [mymax,~] = max(obj.data(:,obj.neckMy));
                    maxVal{index,1} = 'NeckMy (Nm)';
                    maxVal{index,2} = mymax;
                    index = index + 1;
                else
                    %mymax = 0;
                end
                
                % NeckMz max
                if obj.neckMz > 0
                    [mzmax,~] = max(obj.data(:,obj.neckMz));
                    maxVal{index,1} = 'NeckMz (Nm)';
                    maxVal{index,2} = mzmax;
                    index = index + 1;
                else
                    %mzmax = 0;
                end
                
                % Neck Moment Resultant max
                if mres > 0
                    [mresmax,~] = max(mres);
                    maxVal{index,1} = 'Neck Moment Resultant (Nm)';
                    maxVal{index,2} = mresmax;
                    index = index + 1;  %#ok<NASGU>
                else
                    %mresmax = 0;
                end
                
                % Add to object
                obj.maxVals = maxVal;
                
                disp('Finished calculating Max values');
         end
         
         %% Calculates the minimum values in the data
               
         function calcMinVals(obj)
                %%% MINIMUM VALUES
                disp('Calcualting min values...');
                
                global ares fres mres;
                [sledxmin,~] = min(obj.data(:,obj.sledG));
                [axmin,~] = min(obj.data(:,obj.headGx));
                [aymin,~] = min(obj.data(:,obj.headGy));
                [azmin,~] = min(obj.data(:,obj.headGz));
                [aresmin,~] = min(ares);

                
                [fxmin,~] = min(obj.data(:,obj.neckFx));
                if obj.neckFy > 0
                    [fymin,~] = min(obj.data(:,obj.neckFy));
                else
                    fymin = 0;
                end
                [fzmin,~] = min(obj.data(:,obj.neckFz));
                [fresmin,~] = min(fres);

                
                if obj.neckMx > 0
                    mxmin = min(obj.data(:,obj.neckMx));
                else
                    mxmin = 0;
                end
                if obj.neckMy > 0
                    [mymin,~] = min(obj.data(:,obj.neckMy));
                else
                    mymin = 0;
                end
                if obj.neckMz > 0
                    [mzmin,~] = min(obj.data(:,obj.neckMz));
                else
                    mzmin = 0;
                end
                [mresmin,~] = min(mres);

                obj.minVals = [sledxmin axmin aymin azmin aresmin fxmin ...
                               fymin fzmin fresmin mxmin mymin mzmin mresmin];
                
        end
          
            %% Alows changing of data limits for graphs
            function changeLimits(obj, graphStart, graphEnd)
                
                % Need to convert from times to counts
                
                cntsGraphStart = (graphStart/1000) * obj.sampleFreq;
                cntsGraphEnd = (graphEnd/1000) * obj.sampleFreq;
                
                obj.setcutFront(cntsGraphStart);
                obj.setcutRear(cntsGraphEnd);
                %fprintf ('Upper Limit(cnt): %d, Lower limit(cnt): %d\n', obj.cutStart, obj.cutEnd);
            end
            
            %% These functions allow changing of class properties
            function setcutFront(obj, value)
                obj.cutStart = value;
            end
            function setcutRear(obj, value)
                obj.cutEnd = value;
            end
            
            %% Creates a report in the save folder
            function CreateReporttxt(obj)
                % Variables
                picSavepath = obj.plotVars(1);  % file path for save file
                
                % Create and open file
                filename = [char(picSavepath),'\',char(obj.TestId),'\',char(obj.TestId), '_Report.txt' ];
                fid = fopen(filename, 'wt');
                if fid < 0
                    warning('File failed to open')
                
                else
                    % Print the report to file
                    fprintf(fid,'HIC15 is: %.3f\n', obj.HICVal15);
                    fprintf(fid,'HIC36 is: %.3f\n', obj.HICVal36);
                    fprintf(fid,'NIC is: %.3f\n', obj.NICVal);
                    fclose(fid);    % Dont forget to close it =-)
                end
            end
            
            %% Create word document report
            function CreateReportWord(obj)
                
                % variables
                docSavePath = obj.plotVars(1);
                
                % Setup word document
                % Document
                word = actxserver('Word.Application');
                %word.Visible = 1;   % Make word visible for debugging
                document = word.Documents.Add;  % Adds the document
                % Font
                selection = word.Selection;     % Create the selection object
                selection.Font.Name = 'Ariel';  % Select font
                selection.Font.Size = 11;       % Select font size
                
                % Heading
                % heading - Text style setup
                selection.Font.Size = 20;
                selection.ParagraphFormat.Alignment = 1;
                selection.Font.Bold = 1;
                
                % Heading - Text body
                selection.TypeText('Chest Clip Study');
                selection.TypeParagraph;
                selection.TypeText('Test Number: CCA - 0000X Report');
                selection.TypeParagraph;
                
                
                
%with this command we change the Format
%other formats: 0->Left-aligned
%               1->Center-aligned
%               2->Right-aligned
%               3->Fully justified
%               4->Paragraph characters are distributed to fill the entire
%                  width of the paragraph
%               5->Justified with a medium character compression ratio
%               7->Justified with a high character compression ratio
%               8->Justified with a low character compression ratio
%               9->Justified according to Thai formatting Layout
                
                % Matlab Setup
                % Matlab setup - text style setup
                selection.ParagraphFormat.Alignment = 3;
                selection.Font.Size = 11;
                selection.Font.Bold = 1;
                
                % Matlab setup - Text body
                selection.TypeText('Matlab Setup'); 
                selection.TypeParagraph;
                selection.Font.Bold = 0;
                selection.TypeText('Filename: ');
                selection.TypeParagraph;
                selection.TypeText('Channels: ');
                selection.TypeParagraph;
                selection.TypeText('Upper limit:   Lower Limit: ');
                selection.TypeParagraph;
                selection.TypeText('Smoothing Type: ');
                selection.TypeParagraph;
                selection.TypeText('Smoothing span: ');
                selection.TypeParagraph;

                % Head Injury Criteria
                % HIC - text style setup
                selection.ParagraphFormat.Alignment = 3;
                selection.Font.Size = 11;
                selection.Font.Bold = 1;
                
                % HIC - Text body
                selection.TypeText('HIC');
                selection.Font.Bold = 0;
                selection.TypeParagraph;
                selection.TypeText('HIC15: ');
                selection.TypeParagraph;
                selection.TypeText('HIC36: ');
                selection.TypeParagraph;
                
                % Neck Injury Criteria
                % NIC - text style setup
                selection.ParagraphFormat.Alignment = 3;
                selection.Font.Size = 11;
                selection.Font.Bold = 1;
                
                % NIC - Text body
                selection.TypeText('NIC'); 
                selection.TypeParagraph;
                selection.Font.Bold = 0;
                
                % Max values
                % Max values - text style setup
                selection.ParagraphFormat.Alignment = 3;
                selection.Font.Size = 11;
                selection.Font.Bold = 1;
                
                % Max values - Text body
                selection.TypeText('Max Values');
                selection.Font.Bold = 0;
                selection.TypeParagraph;
                selection.TypeText('Chan 1: ');
                selection.TypeParagraph;
                selection.TypeText('Chan 2: ');
                selection.TypeParagraph;
                
                % Min values
                % Min values - text style setup
                selection.ParagraphFormat.Alignment = 3;
                selection.Font.Size = 11;
                selection.Font.Bold = 1;
                
                % Min values - Text body
                selection.TypeText('Min Values'); 
                selection.TypeParagraph;
                selection.Font.Bold = 0;
                selection.TypeText('Chan 1: ');
                selection.TypeParagraph;
                selection.TypeText('Chan 2: ');
                selection.TypeParagraph;
                
                
                % import figures here
              
                % Save the document - dose not create the folder, yet.
                document.SaveAs2([char(docSavePath) '\testWord.docx' ] );
                
                wordQuit = 1;
                if wordQuit == 1
                    word.Quit();    % Debugger can close word
                end
            end
    end
    
    %%
    methods (Static)
    
    end
end
