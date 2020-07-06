# HCVcolocalization
Automated image analysis of cells infected with HCV, measuring co-localization of lipid droplets, core, and ER

Analysis of 3D confocal stacks of HCV infected cells

The program performs the following operations:
   - Automatic threshold estimation of 4 channels
   - Colocalization analysis of channels 2 and 1, 2 and 3
   - Surface detection on all channels
   - Colocalization analysis using volumes of channels 2 and 1, 2 and 3
   - Extraction of statistics
   - Export of channels and merged channel

 Developed on Matlab 2018b (9.4) and  Imaris 9.4.
 Requires: 
 - Apache POI 3.8 (Java Class Library)
 - EasyXT (Matlab-Imaris library)
 - inputsdlg (Matlab extension)
 - xlwrite (Matlab extenstion)

           part of the code was derived from the
           XT_MJG_Surface_Surface_coloc ImarisXT plugin

 INPUTS AS FUNCTION ARGUMENTS
 analysisFolder (string): folder containing files and/or folders to be analyzed

 OUTPUTS AS FILES
 Excel spreadsheet of the tabulated results
 Image file
