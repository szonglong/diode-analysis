# Set of Python files to process Solar Simulator and SCLC data
-----------------------------------------------------------------

Last updated: 01/08/2018



pytest.py, Diode_analysis.cmd
===============================================================
Make graphs and a concat .xls from Keithley voltage sweep data
Author: Zong Long 16/07/2018
Version: 1.9.2
Last updated: 01/08/2018

Description
Running this cmd brings up a prompt to select folder where DATA is kept at. After selection, the python file runs to produce graphs and a concat + analysis (for bright only) .xls file. To edit, both the .cmd file and the .py file will have to be edited in harmony. Note that file directory for .py is hard coded at the moment.

Before running: Make sure that 
1) Anaconda or equivalent is installed in user path variable, or all modules are installed previously.
2) the .cmd file references the correct .py path

Alternatively:
From Spyder or any other IDLE, run .py file, assuming modules are installed correctly. These modules include pandas, numpy, scipy, Tkinter, tkFileDialog



dytest.py
===============================================================
Does SCLC analysis to calculate mobility
Author: Zong Long 20/07/2018
Version: 2.0
Last updated: 26/07/2018

Description
Plots a semi-log and log/log for front and back voltage sweeps for SCLC analysis (typical +- 8V) and produces a .xls file containing: concat, analysis, and derivatives. It selects the region where a user defined list_tol number of pixels are in SCLC regime (exp = 2) and mobility.



mtest.py
===============================================================
Plots smoothed local ideality curves for dark graphs
Author: Zong Long 30/07/2018
Version: 1.2
Last updated: 01/08/2018

Description
Helps the user find the exponential regime by first plotting a semi-log curve and also a local ideality curve. This curve is doubly smooth using the savgol algorithm once on the current curve and again after the ideality curve is plotted. Only concerned about first sweep from 0-3V.






###############################################################

Known bugs:
-First pixel must always be a complete sweep as length of df is based on that


---------------------
Changelog can be found on github/szonglong/diode-analysis




