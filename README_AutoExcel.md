How to use the scripts:
<b>AutoExcelCollation_basic.m</b><br>
<b>AutoExcelCollation_Bubble_IQR.m</b>

Both scripts are operated identically. <br>
The first script will collate all the 3DEP results within a single folder, group them by frequency, and generate a three column summary table [freq, avg, sdDEV]. The first two columns of which are easily copy/pasted into the semiAutofit_MPJ_v#.m code also on this Github. The table is dynamic, removing values from Column B will update the table, do not delete rows or the automatic equations will likely break. <br>
The later script ('...._Bubble_IQR') contains additional filtering to automatically remove datapoints the 3DEP flagged as bubbles, and automatically generates an interquartile range test for each frequencies datapoints and removes outliers. In both cases the exceptional datapoint values are printed in column E and can easily be dragged back into column B if there was a mistake and they are desired to be included in further analysis.

<b>Operation</b>

A) Make sure the 3DEP repeats are in a single folder on your PC, it is recommended to end each sample name with A,B,C... or 1,2,3... for example:
<img width="500" height="400" alt="image" src="https://github.com/user-attachments/assets/18d893cc-cd9e-4e93-ae8e-2e1ecd902178" />

B) Run the code and identify the folder on your PC in the popup, notice that the folder will likely appear empty.
<img width="500" height="350" alt="image" src="https://github.com/user-attachments/assets/a3c89e0f-94f1-4bf6-a2d7-e429d88cefde" />

C) The command window will tell you when the operation is complete. The excel can then be reopened and appears as the below example.
<img width="1000" height="350" alt="image" src="https://github.com/user-attachments/assets/45c4067d-b2d7-4002-8a6e-44d80b2f097f" />
<img width="500" height="500" alt="image" src="https://github.com/user-attachments/assets/609b90e9-33c6-4041-a197-1dcf90f56043" />

<b> Output format: </b><br>
Column A: Datapoint Frequency <br>
Column B: Datapoint Relative Polarisability (raw 3DEP output) <br>
Column C: This column contains a string consisting of the final digit in the sample's .csv filename, followed immediately by any comment the 3DEP automatically attached to that datapoint such as 'Bubble' or 'Outlier'. <br>
Column D: If datapoints were removed by the IQR or Bubble filters then the values are saved in this column. <br>
Column F: Can be ignored. This is used by the equations behind the scenes to identify the row numbers for the start and end of each frequency. <br>
Column G: Summary table frequencies <br>
Column H: Summary table polarisability average values <br>
Column I: Summary table polarisability stdDEv values <br>
