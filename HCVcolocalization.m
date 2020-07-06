function HCVcolocalization(varargin)
% Revision 6
% Analysis of 3D confocal stacks of HCV infected cells
%
% The program performs the following operations:
%    - Automatic threshold estimation of 4 channels
%    - Colocalization analysis of channels 2 and 1, 2 and 3
%    - Surface detection on all channels
%    - Colocalization analysis using volumes of channels 2 and 1, 2 and 3
%    - Extraction of statistics
%    - Export of channels and merged channel
%
% Developed on Matlab 2017b (9.3) and  Imaris 9.1.2
% Requires: Apache POI 3.8 (Java Class Library)
%           EasyXT (Matlab-Imaris library)
%           inputsdlg (Matlab extension)
%           xlwrite (Matlab extenstion)
%
%           part of the code was derived from the
%           XT_MJG_Surface_Surface_coloc ImarisXT plugin
%
% INPUTS AS FUNCTION ARGUMENTS
% analysisFolder (string): folder containing files and/or folders to be analyzed
%
% OUTPUTS AS FILES
% Excel spreadsheet of the tabulated results
% Image files
%
% (c) Andrea Galli, CO-HEP, Copenhagen, 2018-20

Parameters = struct('analysisFolder', {''},...
                    'baseFolder', {''},...
                    'baseFilename', {''},...
                    'fileName', {''},...
                    'doThresholds',false,...
                    'doExtraction',false,...
                    'doColocalization',false,...
                    'doSurfaces',false,...
                    'doVolumeColocalization',false,...
                    'doStatistics',false,...
                    'thresholdPercent',10,...
                    'channel',{...
                           struct('name',{'ER', 'Core', 'LDs', 'Nucleus'},...
                                  'color', {255, 65535, 65280, 16711680},...
                                  'threshold',{0, 0, 0, 0},...
                                  'max',{0, 0, 0, 0},...
                                  'voxels',{0, 0, 0, 0})},...
                    'xlsResultsFilename',{''},...
                    'xlsColocSheet',{'Colocalization'},...
                    'xlsStatSheet',{'LD Statistics'},...
                    'xlsVolSheet',{'LD Volumes statistics'},...
                    'xlsSphSheet',{'LD Sphericieties statistics'},...
                    'xlsStartCell',{'A2'});

    if nargin<1
        Parameters.analysisFolder=uigetdir(pwd,'Samples folder selection');
        if ~Parameters.analysisFolder
            return
        end
    else
        Parameters.analysisFolder=varargin{1};
    end

    %% Define Parameters and constants
    if ispc
        backslash='\';
    else
        backslash ='/';
    end

    colorMapRed=[linspace(0,1,256);linspace(0,0,256);linspace(0,0,256)];
    colorMapRed=colorMapRed';
    colorMapGreen=[linspace(0,0,256);linspace(0,1,256);linspace(0,0,256)];
    colorMapGreen=colorMapGreen';
    colorMapBlue=[linspace(0,0,256);linspace(0,0,256);linspace(0,1,256)];
    colorMapBlue=colorMapBlue';
    colorMapYellow=[linspace(0,1,256);linspace(0,1,256);linspace(0,0,256)];
    colorMapYellow=colorMapYellow';
    colorMaps={colorMapRed,colorMapYellow,colorMapGreen,colorMapBlue};

    %% Collect user-defined Parameters
    [exitCode,Parameters] = GetAnalysisParameters(Parameters);
    if ~exitCode
        return
    end

    %% Collect file names
    [filesList,nFiles]=GetFiles(Parameters.analysisFolder);

    %% Data Analysis
    % Define an xls name for results
    if Parameters.doColocalization == 1 || Parameters.doStatistics == 1
        [resultsFile, resultsFolder] = uiputfile({'*.xls,*.xlsx', 'Excel files (*.xls, *.xlsx)'}, 'Save results as',...
                                                 strcat(Parameters.analysisFolder,backslash,'Results.xlsx'));
        Parameters.xlsResultsFilename = strcat(resultsFolder, backslash, resultsFile);
        colocHeader = {'Group' 'Sample' 'Core/ER' 'Core/LD' 'ER max' 'ER thr' 'Core max' 'Core thr' 'LD max' 'LD thr'};
        xlwrite(Parameters.xlsResultsFilename, colocHeader, Parameters.xlsColocSheet,'A1');
        statHeader = {'Group' 'Sample' 'LD No' 'LD Vol Min' 'LD Vol Max' 'LD Vol Median' 'LD Vol Total' 'LD Vol SD' ...
                       'LD Sphericity Avg' 'LD Sphericity SD' 'ER Vol total' 'ER Vol SD' 'Core Vol Total' 'Core Vol SD'...
                       'Nucleus Vol Total' 'Nucleus Vol SD' 'Surface Coloc Core/LD' 'Surface Coloc Core/ER'};
        xlwrite(Parameters.xlsResultsFilename, statHeader, Parameters.xlsStatSheet,'A1');
    end

    clc;
    tic;
    fprintf('Beginning analysis\n');

    statResults = cell(nFiles,18);
    volResults = cell(nFiles,1000);
    sphResults = cell(nFiles,1000);
    colocResults = cell(nFiles,10);
    for iFile=1:nFiles
        [~,Parameters.baseFilename,~]=fileparts(filesList{iFile,2});        %Obtain the base filename and extension
        pathParts=regexp(filesList{iFile,1},backslash,'split');             %Obtain the directory containing the file
        Parameters.baseFolder=pathParts{size(pathParts,2)};
        Parameters.fileName = strcat(filesList(iFile,1), backslash, filesList(iFile,2));
        ImarisXT = EasyXT(StartImaris);
        fprintf('\nAnalyzing file %i of %i: %s\n', iFile, nFiles, Parameters.baseFilename);
        ImarisXT.OpenImage(Parameters.fileName);
        %% Calculate Thresholds
        if Parameters.doThresholds
            fprintf('   Thresholding...');
            Parameters = CalculateThresholds(ImarisXT, Parameters);
            fprintf('done!\n');
        end

        %% Colocalization Analysis
        if Parameters.doColocalization
            fprintf('   Calculating colocalization...');
            colocResults(iFile,:) = ColocalizationAnalysis (ImarisXT, Parameters); %add results to the Results cell array
            xlwrite(Parameters.xlsResultsFilename, colocResults, Parameters.xlsColocSheet, Parameters.xlsStartCell);
            fprintf('done!\n');
        end

        %% Detect surfaces
        if Parameters.doSurfaces
            fprintf('   Detecting surfaces...');
            VolumeAnalysis (ImarisXT, Parameters);
            fprintf('done!\n');
        end

        %% Perform volume colocalization analysis
        if Parameters.doVolumeColocalization
            fprintf('   Calculating volumetric colocalization...');
            % delete all colocalized volumes
            imarisScene = ImarisXT.ImarisApp.GetSurpassScene;
            nSurfaces = 0;
            iObject=0;
            while iObject < imarisScene.GetNumberOfChildren
                imarisObject = imarisScene.GetChild(iObject);
                if ImarisXT.ImarisApp.GetFactory.IsSurfaces(imarisObject)
                    if nSurfaces < 4
                        nSurfaces = nSurfaces+1;
                        iObject = iObject+1;
                    else
                        imarisScene.RemoveChild(imarisObject);
                    end
                else
                    iObject = iObject+1;
                end
            end
            VolumeColocalization(ImarisXT, [2 3], Parameters);
            VolumeColocalization(ImarisXT, [2 1], Parameters);
            fprintf('done!\n');
        end

        %% Export volume statistics
        if Parameters.doStatistics && ImarisXT.GetNumberOf('Surfaces')>5
            fprintf('   Exporting volume statistics...');
            [newStats, newVolStats, newSphResults] = ExtractStatistics(ImarisXT, Parameters);
            volResults(iFile,1:(size(newVolStats,2)))= newVolStats;
            sphResults(iFile,1:(size(newSphResults,2)))= newSphResults;
            statResults(iFile,:) = newStats; %add results to the Results cell array
            xlwrite(Parameters.xlsResultsFilename, statResults, Parameters.xlsStatSheet, Parameters.xlsStartCell);
            xlwrite(Parameters.xlsResultsFilename, volResults', Parameters.xlsVolSheet, 'A1');
            xlwrite(Parameters.xlsResultsFilename, sphResults', Parameters.xlsSphSheet, 'A1');
            fprintf('done!\n');
        end

        %% Export single channels and merged image
        if Parameters.doExtraction
            fprintf('   Exporting channels...');
            allChannels=zeros(1024,1024,3,5);
            for iChannel=1:5
                if iChannel<5
                    channelImage=int8(ImarisXT.GetVoxels(iChannel));
                    [Parameters.channel(iChannel).threshold,~] = ImarisXT.GetChannelRange(iChannel);
                    MIP=gray2ind(imadjust(mat2gray((max(channelImage,[],3) -  Parameters.channel(iChannel).threshold),[0 255])),256);
                    allChannels(:,:,:,iChannel)=ind2rgb(MIP, colorMaps{iChannel});
                    fullImageFilename = strcat(filesList(iFile,1),backslash,Parameters.baseFilename, '_',Parameters.channel(iChannel).name,'.tif');
                    imwrite(allChannels(:,:,:,iChannel), fullImageFilename{1},'tif','compression','lzw');
                else
                    allChannels(:,:,:,iChannel)=max(allChannels,[],4); %Merge channels
                    fullImageFilename = strcat(filesList(iFile,1),backslash,Parameters.baseFilename, '_Merged','.tif');
                    imwrite(allChannels(:,:,:,iChannel), fullImageFilename{1},'tif','compression','lzw');
                end
            end
            fprintf('done!\n');
        end
        ImarisXT.ImarisApp.FileSave(Parameters.fileName,'writer="Imaris5"');
        ImarisXT.ImarisApp.Quit;
        fprintf('Complete\n');
    end
    fprintf(1,'Analysis completed in %2.2f minutes \n', toc/60 );
end
%% Analysis functions
function returnParameters = CalculateThresholds (ImarisXT, Parameters)
    % Load channels
    for iChannel=1:4
        Parameters.channel(iChannel).voxels = ImarisXT.GetVoxels(iChannel,'1D',1);
        Parameters.channel(iChannel).max = double(max(Parameters.channel(iChannel).voxels));
    end
    % Calculate threshold for each channel
    Parameters.channel(1).threshold = Parameters.channel(1).max * Parameters.thresholdPercent / 100;
    Parameters.channel(4).threshold = Parameters.channel(4).max * Parameters.thresholdPercent / 100;
    PairedCh2Ch3=[Parameters.channel(2).voxels Parameters.channel(3).voxels];
    overallMaximum = double(max(max(PairedCh2Ch3)));
    if (overallMaximum <= 255)
        overallMaximum = 255;
    end
    centers{1} = linspace(0,overallMaximum,256);
    centers{2} = linspace(0,overallMaximum,256);
    BimodalCh2Ch3 = hist3(PairedCh2Ch3,centers);
    BimodalCh2Ch3Thr=BimodalCh2Ch3>0;
    Temp=sum(BimodalCh2Ch3Thr,2);
    Temp=Temp';
    Max2=FindPeak(Temp);
    %[col,Max2]=max(sum(M,2));
    Parameters.channel(2).threshold=2*Max2;
    %[~,Max3]=max(sum(M,1));
    Max3=FindPeak(sum(BimodalCh2Ch3Thr,1));
    Parameters.channel(3).threshold=2*Max3;
    % Write data into Imaris file
    for iChannel=1:4
        ImarisXT.SetChannelRange(iChannel, Parameters.channel(iChannel).threshold, Parameters.channel(iChannel).max);
        ImarisXT.SetChannelName(iChannel, Parameters.channel(iChannel).name);
        ImarisXT.SetChannelColor(iChannel, Parameters.channel(iChannel).color);
    end
    returnParameters=Parameters;
end
function colocalizationData = ColocalizationAnalysis (ImarisXT, Parameters)
    % Read channels info
    for iChannel=1:4
        [rangeMin,rangeMax]= ImarisXT.GetChannelRange(iChannel);
        Parameters.channel(iChannel).threshold = rangeMin;
        Parameters.channel(iChannel).max = rangeMax;
    end
    % Remove preexisting extra channels
    if ImarisXT.GetSize('C')>4
        ImarisXT.SetSize('C',4);
    end
    colocalizationCoefficients = ImarisXT.Coloc([2 1; ...
                                         2 3],...
                                        [Parameters.channel(2).threshold Parameters.channel(1).threshold;...
                                         Parameters.channel(2).threshold Parameters.channel(3).threshold],...
                                        'Coloc Channel', true);
    colocalizationData{1,1} = Parameters.baseFolder;
    colocalizationData{1,2} = Parameters.baseFilename;
    colocalizationData{1,3} = colocalizationCoefficients.matA(1);
    colocalizationData{1,4} = colocalizationCoefficients.matA(2);
    colocalizationData{1,5} = Parameters.channel(1).max;
    colocalizationData{1,6} = Parameters.channel(1).threshold;
    colocalizationData{1,7} = Parameters.channel(2).max;
    colocalizationData{1,8} = Parameters.channel(2).threshold;
    colocalizationData{1,9} = Parameters.channel(3).max;
    colocalizationData{1,10}= Parameters.channel(3).threshold;
end
function VolumeAnalysis (ImarisXT, Parameters)
    % Removes all pre-existing surfaces
    for iSurface=1:ImarisXT.GetNumberOf('Surfaces')
        temporaryImarisObject=ImarisXT.GetObject('Type', 'Surfaces','Number', 1);
        ImarisXT.RemoveFromScene(temporaryImarisObject);
    end
    %detect new surfaces from channels 1 to 4
    for iChannel=1:4
        [~, rangeMax]=ImarisXT.GetChannelRange(iChannel);
        newSurface = ImarisXT.DetectSurfaces(iChannel, 'Name', Parameters.channel(iChannel).name,...
                                                       'Threshold', rangeMax*0.15, ...
                                                       'Color', Parameters.channel(iChannel).color, ...
                                                       'Smoothing', 0.0930);
        ImarisXT.AddToScene(newSurface);
    end
end
function VolumeColocalization(ImarisXT, channelSurfaces, Parameters)
    %clone Dataset
    imarisApplication = ImarisXT.ImarisApp;
    workDataset = imarisApplication.GetDataSet.Clone;

    % get all Surpass surfaces
    imarisScene = imarisApplication.GetSurpassScene;

    nSurfaces = 0;
    nObjects=imarisScene.GetNumberOfChildren;

    surfacesList{nObjects} = [];
    for iObject = 1:nObjects
        imarisObject = imarisScene.GetChild(iObject - 1);
        if imarisApplication.GetFactory.IsSurfaces(imarisObject)
            nSurfaces = nSurfaces+1;
            surfacesList{nSurfaces} = imarisApplication.GetFactory.ToSurfaces(imarisObject);
        end
    end

    %Choose the surfaces
    surface1 = surfacesList{channelSurfaces(1)};
    surface2 = surfacesList{channelSurfaces(2)};

    %Get Image Data parameters
    [datasize, extentMin, extentMax]=ImarisXT.GetExtents();
    maxX = extentMax(1);
    maxY = extentMax(2);
    maxZ = extentMax(3);
    minX = extentMin(1);
    minY = extentMin(2);
    minZ = extentMin(3);
    sizeX = datasize(1);
    sizeY = datasize(2);
    sizeZ = datasize(3);

    nChannels = ImarisXT.GetSize('C');
    voxelSize= (maxX-minX)/sizeX;
    smoothingFactor=voxelSize*2;

    %add additional channel
    workDataset.SetSizeC(nChannels + 1);
    lastChannel=nChannels;
    time = 0;
    %Generate surface mask for each surface
    surface1Mask = surface1.GetMask(minX,minY,minZ,maxX,maxY,maxZ,sizeX, sizeY,sizeZ,time);
    surface2Mask = surface2.GetMask(minX,minY,minZ,maxX,maxY,maxZ,sizeX, sizeY,sizeZ,time);

    maskingChannel1 = surface1Mask.GetDataVolumeAs1DArrayBytes(0,time);
    maskingChannel2 = surface2Mask.GetDataVolumeAs1DArrayBytes(0,time);

    %Determine the Voxels that are colocalized
    colocalizationChannel=maskingChannel1+maskingChannel2;
    colocalizationChannel(colocalizationChannel<2)=0;
    colocalizationChannel(colocalizationChannel>1)=1;

    surfaceName= sprintf('Volume Colocalization %s-%s', Parameters.channel(channelSurfaces(1)).name, Parameters.channel(channelSurfaces(2)).name);
    workDataset.SetDataVolumeAs1DArrayBytes(colocalizationChannel, lastChannel, time);
    workDataset.SetChannelName(lastChannel,surfaceName);
    workDataset.SetChannelRange(lastChannel,0,1);
    imarisApplication.SetDataSet(workDataset);
    %Run the Surface Creation Wizard on the new channel
    ip = imarisApplication.GetImageProcessing;
    colocalizationSurface = ip.DetectSurfaces(workDataset, [], lastChannel, smoothingFactor, 0, true, 55, '');
    colocalizationSurface.SetName(surfaceName);
    colocalizationSurface.SetColorRGBA((rand(1, 1)) * 256 * 256 * 256 );
    %Add new surface to Surpass Scene
    imarisApplication.GetSurpassScene.AddChild(colocalizationSurface, -1);
    ImarisXT.SetSize('C',lastChannel);
end
%% Service functions
function [exitCode, updatedParameters] = GetAnalysisParameters(Parameters)
    % Define option dialog parameters
    dialogTitle = 'Image analysis options';
    % SETTING DIALOG OPTIONS
    Options.WindowStyle = 'modal';
    Options.Resize = 'on';
    Options.Interpreter = 'tex';
    Options.CancelButton = 'on';
    Options.ApplyButton = 'off';
    Options.ButtonNames = {'OK','Cancel'};

    dialogFormats = {};
    dialogDefAns = struct([]);

    dialogPrompt = {'Intensity Analysis: Thresholding' 'CalcThresholds',[]
                    'Colocalization' 'CalcColoc',[]
                    'Percentage threshold', 'PercThreshold',[]
                    'Volume Analysis: Detection' 'CalcSurfaces',[]
                    'Colocalization' 'CalcVolColoc',[]
                    'Extract volume statistics' 'CalcStatistics',[]
                    'Extract single-channel images' 'CalcImages',[]
                    };

    dialogFormats(1,1).type = 'check';
    dialogFormats(1,2).type = 'check';
    dialogFormats(1,3).type = 'edit';
    dialogFormats(1,3).format = 'integer';
    dialogFormats(1,3).limits = [0 100];
    dialogFormats(1,3).size = 25;
    dialogFormats(1,3).unitsloc = 'bottomleft';
    dialogFormats(2,1).type = 'check';
    dialogFormats(2,2).type = 'check';
    dialogFormats(3,1).type = 'check';
    dialogFormats(4,1).type = 'check';

    dialogDefAns(1).CalcThresholds = Parameters.doThresholds;
    dialogDefAns.CalcColoc = Parameters.doColocalization;
    dialogDefAns.CalcSurfaces = Parameters.doSurfaces;
    dialogDefAns.CalcVolColoc = Parameters.doVolumeColocalization;
    dialogDefAns.CalcStatistics = Parameters.doStatistics;
    dialogDefAns.PercThreshold = Parameters.thresholdPercent;
    dialogDefAns.CalcImages = Parameters.doExtraction;

    [Answers,Cancelled] = inputsdlg(dialogPrompt,dialogTitle,dialogFormats,dialogDefAns,Options);
    if not(Cancelled)
        Parameters.doThresholds = Answers.CalcThresholds;
        Parameters.doExtraction = Answers.CalcImages;
        Parameters.doColocalization = Answers.CalcColoc;
        Parameters.doSurfaces = Answers.CalcSurfaces;
        Parameters.doStatistics = Answers.CalcStatistics;
        Parameters.doVolumeColocalization = Answers.CalcVolColoc;
        Parameters.thresholdPercent = double(Answers.PercThreshold);
        exitCode=true;
    else
        exitCode=false;
    end
    updatedParameters=Parameters;
end
function [filesList,nFiles] = GetFiles (targetFolder)
    %% Collect file names from a folder and its subfolders
    % Read files
    nFiles=0;
    nGroups=0;
    filesList=cell(500,3);
    startFolder = pwd;
    cd(targetFolder);
    folderContent=dir('*.ims');
    nItems=length(folderContent);
    if nItems>0
        nGroups=nGroups+1;
        for iItem=1:nItems
            nFiles=nFiles+1;
            filesList(nFiles,1)={cd};
            filesList(nFiles,2)={folderContent(iItem).name};
            filesList(nFiles,3)={nGroups};
        end
    end
    % Read subfolders
    folderContent = dir;
    nItems=length(folderContent);
    for iItem=1:nItems
        if folderContent(iItem).isdir == 1
            if folderContent(iItem).name == '.'
            else
                cd(folderContent(iItem).name);
                SubFolderContent=dir('*.ims');
                SubFolderItems=length(SubFolderContent);
                if SubFolderItems>0
                    nGroups=nGroups+1;
                    for iSubItem=1:SubFolderItems
                        nFiles=nFiles+1;
                        filesList(nFiles,1)={cd};
                        filesList(nFiles,2)={SubFolderContent(iSubItem).name};
                        filesList(nFiles,3)={nGroups};
                    end
                end
                cd ('../');
            end
        end
    end
    cd(startFolder);
    filesList(all(cellfun(@isempty,filesList),2),:)=[]; %remove empty cells from FileList
    [nFiles,~]=size(filesList);                                           %How many files?
end
function [statData,volData,sphData] = ExtractStatistics (ImarisXT, Parameters)
    temporaryImarisObject=ImarisXT.GetObject('Name', 'LDs');
    statLdNum=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Total Number of Surfaces');
    statLdVol=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Volume');
    statLdSph=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Sphericity');
    ldVolumes=statLdVol.values;
    lsSphericities=statLdSph.values;
    temporaryImarisObject=ImarisXT.GetObject('Name', 'ER');
    statErVol=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Volume');
    erVolumes=statErVol.values;
    temporaryImarisObject=ImarisXT.GetObject('Name', 'Core');
    statCoreVol=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Volume');
    coreVolumes=statCoreVol.values;
    temporaryImarisObject=ImarisXT.GetObject('Name', 'Nucleus');
    statNucVol=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Volume');
    nucleusVolumes=statNucVol.values;
    temporaryImarisObject=ImarisXT.GetObject('Name', 'Volume Colocalization Core-LDs');
    statColocLdVol=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Volume');
    colocLdVolumes=statColocLdVol.values;
    temporaryImarisObject=ImarisXT.GetObject('Name', 'Volume Colocalization Core-ER');
    statColocERVol=ImarisXT.GetSelectedStatistics(temporaryImarisObject, 'Volume');
    colocErVolumes=statColocERVol.values;
    statData{1,1}=Parameters.baseFolder;
    statData{1,2}=Parameters.baseFilename;
    statData{1,3}=statLdNum.values;
    statData{1,4}=min(ldVolumes);
    statData{1,5}=max(ldVolumes);
    statData{1,6}=median(ldVolumes);
    statData{1,7}=sum(ldVolumes);
    statData{1,8}=std(ldVolumes);
    statData{1,9}=mean(lsSphericities);
    statData{1,10}=std(lsSphericities);
    statData{1,11}=sum(erVolumes);
    statData{1,12}=std(erVolumes);
    statData{1,13}=sum(coreVolumes);
    statData{1,14}=std(coreVolumes);
    statData{1,15}=sum(nucleusVolumes);
    statData{1,16}=std(nucleusVolumes);
    statData{1,17}=sum(colocLdVolumes);
    statData{1,18}=sum(colocErVolumes);
    %extract single Volume values
    volData{1,1}=Parameters.baseFolder;
    volData{1,2}=Parameters.baseFilename;
    volData=[volData,num2cell(ldVolumes',1)];
    %extract single Sphericity values
    sphData{1,1}=Parameters.baseFolder;
    sphData{1,2}=Parameters.baseFilename;
    sphData=[sphData,num2cell(lsSphericities',1)];

end
%% Accessory functions
function [firstPeakIndex] = FindPeak(intensityHistogram)
    extendedHistogram = [ 0 0 0 0 0 0 0 0 0 0 0 0 intensityHistogram 0 0 0 0 0 0 0 0 0 0 0 0];
    for i=1:length(extendedHistogram)-12
        slidingWindow=extendedHistogram(i:i+24);
        if(max(slidingWindow)==slidingWindow(13))
            firstPeakIndex = i;
            return;
        end
    end
end
function imarisApplication = StartImaris ()
    %% Connection to Imaris
        fprintf('\nStarting Imaris...');
        %Say('Starting Imaris');
        if ispc
			!'C:\Program Files\Bitplane\Imaris x64 9.1.2\Imaris.exe' &
        else
			!'/Applications/Imaris 9.2.1.app/Contents/MacOS/Imaris' &
        end
        pause(7);
        imarisLib = ImarisLib;
        imarisServer = imarisLib.GetServer;
        try
            % Connect to the last opened Imaris instance
            nImarisObjects=imarisServer.GetNumberOfObjects;
            lastObjectID=imarisServer.GetObjectID(nImarisObjects-1);
            imarisApplication = imarisLib.GetApplication(lastObjectID);
            catch err
                Say('Houston we have a problem');
                disp('Could not get Imaris Application with ID');
                disp('Original error below:');
                rethrow(err);
        end
        fprintf('connection complete\n');
end
function Say (yourThoughts)
    if ispc
        SV = actxserver('SAPI.SpVoice');
        invoke(SV,'Speak',yourThoughts);
        delete(SV);
        clear SV;
        pause(0.2);
    elseif ismac
        system(sprintf('say %s &', yourThoughts));
    else
    end
end
