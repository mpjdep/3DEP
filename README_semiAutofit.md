How to use the script:
<b>semiAutofit_MPJ_v#.m</b>

A) line 18 needs to be updated with the average radius of the cells.
B) line 19 needs to be updates with the conductivity of the solution, in S/m.

<img width="1883" height="137" alt="image" src="https://github.com/user-attachments/assets/ef2c8a1c-cbd6-4800-934f-93b6f5f5f9eb" />

C) Copy and paste your 3DEP data (first column frequency, second column relative polarisability) into the square brackets starting on line 22.

<img width="382" height="395" alt="image" src="https://github.com/user-attachments/assets/caae7ea6-85f1-418a-8a5e-a4ab87a5be85" />

D) If required you can amend the initial guesses, and minimum/maximum limits for the output variables below the raw data in lines 43-45 (lines can vary).

<img width="840" height="120" alt="image" src="https://github.com/user-attachments/assets/3b78401b-0f28-429c-9ad0-ed7ec6ba2341" />

E) Run the script

<img width="212" height="105" alt="image" src="https://github.com/user-attachments/assets/d33f979b-9033-4eeb-aaf9-16b1765a4c52" /><br>

<img width="350" height="307" alt="image" src="https://github.com/user-attachments/assets/4dca7d91-e14e-4945-a355-8efbeeeb8ac5" />

F) Anomalous datapoints can be removed from the data=[ ...]; section, or commented out, if required.
